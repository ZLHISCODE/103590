VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchaseVerifySelect 
   Caption         =   "财务审核查询"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8610
   Icon            =   "frmPurchaseVerifySelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   8610
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList imgPicture 
      Left            =   4560
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseVerifySelect.frx":6852
            Key             =   "old"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   3840
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5895
      ScaleWidth      =   15
      TabIndex        =   20
      Top             =   120
      Width           =   10
   End
   Begin TabDlg.SSTab sstInfo 
      Height          =   4935
      Left            =   4200
      TabIndex        =   9
      Top             =   240
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "汇总信息"
      TabPicture(0)   =   "frmPurchaseVerifySelect.frx":D0B4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vsfAll"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkALLVisible1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "明细信息"
      TabPicture(1)   =   "frmPurchaseVerifySelect.frx":D0D0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblGroup"
      Tab(1).Control(1)=   "vsfList"
      Tab(1).Control(2)=   "optFloor"
      Tab(1).Control(3)=   "optMedi"
      Tab(1).Control(4)=   "chkALLVisible2"
      Tab(1).ControlCount=   5
      Begin VB.CheckBox chkALLVisible2 
         Caption         =   "显示完整财务审核单据"
         Height          =   180
         Left            =   -74880
         TabIndex        =   19
         Top             =   510
         Width           =   2175
      End
      Begin VB.OptionButton optMedi 
         Caption         =   "药品分组"
         Height          =   180
         Left            =   -69600
         TabIndex        =   18
         Top             =   510
         Width           =   1095
      End
      Begin VB.OptionButton optFloor 
         Caption         =   "单据分组"
         Height          =   180
         Left            =   -70800
         TabIndex        =   17
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox chkALLVisible1 
         Caption         =   "显示完整财务审核单据"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAll 
         Height          =   1845
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "红色字体表示财务审核单据与原始单据或者与上一个单据有区别"
         Top             =   1080
         Width           =   3255
         _cx             =   5741
         _cy             =   3254
         Appearance      =   1
         BorderStyle     =   1
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPurchaseVerifySelect.frx":D0EC
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   1845
         Left            =   -74880
         TabIndex        =   12
         ToolTipText     =   "红色字体表示财务审核单据与原始单据或者与上一个单据有区别"
         Top             =   1080
         Width           =   3255
         _cx             =   5741
         _cy             =   3254
         Appearance      =   1
         BorderStyle     =   1
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchaseVerifySelect.frx":D1E1
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         Caption         =   "分组方式"
         Height          =   180
         Left            =   -71640
         TabIndex        =   16
         Top             =   510
         Width           =   720
      End
   End
   Begin VB.PictureBox picLeft 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4695
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   60
         Width           =   2535
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         Left            =   840
         TabIndex        =   14
         Top             =   1780
         Width           =   2535
      End
      Begin VB.PictureBox picDate 
         BorderStyle     =   0  'None
         Height          =   800
         Left            =   0
         ScaleHeight     =   795
         ScaleWidth      =   3735
         TabIndex        =   4
         Top             =   800
         Width           =   3735
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   300
            Left            =   840
            TabIndex        =   5
            Top             =   540
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   529
            _Version        =   393216
            Format          =   100401152
            CurrentDate     =   41775
         End
         Begin MSComCtl2.DTPicker dtpBeginDate 
            Height          =   300
            Left            =   840
            TabIndex        =   6
            Top             =   120
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   529
            _Version        =   393216
            Format          =   100401152
            CurrentDate     =   41775
         End
         Begin VB.Label lblBeginDate 
            AutoSize        =   -1  'True
            Caption         =   "开始日期"
            Height          =   180
            Left            =   0
            TabIndex        =   8
            Top             =   180
            Width           =   720
         End
         Begin VB.Label lblEndDate 
            AutoSize        =   -1  'True
            Caption         =   "结束日期"
            Height          =   180
            Left            =   0
            TabIndex        =   7
            Top             =   600
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找"
         Height          =   300
         Left            =   2835
         TabIndex        =   3
         Top             =   500
         Width           =   510
      End
      Begin VB.ComboBox cboDate 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   500
         Width           =   2015
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfLeft 
         Height          =   2205
         Left            =   0
         TabIndex        =   11
         Top             =   2520
         Width           =   3375
         _cx             =   5953
         _cy             =   3889
         Appearance      =   1
         BorderStyle     =   1
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
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchaseVerifySelect.frx":D400
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl库房 
         AutoSize        =   -1  'True
         Caption         =   "库    房"
         Height          =   180
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         Caption         =   "NO"
         Height          =   180
         Left            =   0
         TabIndex        =   15
         Top             =   1840
         Width           =   180
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "日    期"
         Height          =   180
         Left            =   0
         TabIndex        =   1
         Top             =   560
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmPurchaseVerifySelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsData As New ADODB.Recordset  '数据集
Private mrsCloneDta As New ADODB.Recordset '克隆数据集
Private mstr当前库房 As Long  '传过来的当前库房
Private mStr库房 As String  '传过来的库房集合
Private mlng库房ID As Long '当前选中库房
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mdatBeginDate As Date    '开始查询时间
Private mdatEndDate As Date    '结束查询时间

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数

Private Const M_INT_售价单位 As Integer = 1
Private Const M_INT_门诊单位 As Integer = 2
Private Const M_INT_住院单位 As Integer = 3
Private Const M_INT_药库单位 As Integer = 4

Private Sub SetControlLocation()
    '设置控件位置
    On Error Resume Next
    
    picLeft.Move 50, 50, txtNo.Left + txtNo.Width, Me.ScaleHeight - 50
    cmdFind.Move cboDate.Left + cboDate.Width, cboDate.Top
    picDate.Left = 0
    LblNO.Move lblDate.Left, txtNo.Top + 60
    vsfLeft.Move 0, txtNo.Top + txtNo.Height + 100, picLeft.Width, picLeft.ScaleHeight - (txtNo.Top + txtNo.Height + 150)
    picSplit.Move picLeft.Left + picLeft.Width, 0, 10, Me.ScaleHeight
    sstInfo.Move picLeft.Left + picLeft.Width, 50, Me.ScaleWidth - picSplit.Left + 30, Me.ScaleHeight - 50
    chkALLVisible1.Move 100, 480
    chkALLVisible2.Move 100, chkALLVisible1.Top
    lblGroup.Top = chkALLVisible1.Top
    optFloor.Top = chkALLVisible1.Top
    optMedi.Top = chkALLVisible1.Top
    vsfAll.Move 100, chkALLVisible1.Top + chkALLVisible1.Height + 50, sstInfo.Width - 100, sstInfo.Height - (chkALLVisible1.Top + chkALLVisible1.Height + 50)
    vsfList.Move 100, chkALLVisible1.Top + chkALLVisible1.Height + 50, sstInfo.Width - 100, sstInfo.Height - (chkALLVisible1.Top + chkALLVisible1.Height + 50)
End Sub

Private Sub cboDate_Click()
    With cboDate
        If .Text = "自定义" Then
            picDate.Visible = True
            txtNo.Top = picDate.Top + picDate.Height + 120
            LblNO.Top = txtNo.Top + 60
            vsfLeft.Top = txtNo.Top + txtNo.Height + 100
            vsfLeft.Height = picLeft.ScaleHeight - (txtNo.Top + txtNo.Height + 100)
        Else
            picDate.Visible = False
            txtNo.Top = picDate.Top + 100
            LblNO.Top = txtNo.Top + 60
            vsfLeft.Top = txtNo.Top + txtNo.Height + 100
            vsfLeft.Height = picLeft.ScaleHeight - (txtNo.Top + txtNo.Height + 100)
        End If
        
        Select Case .Text
            Case "一个月内"
                mdatBeginDate = CDate(Format(DateAdd("M", -1, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = Sys.Currentdate
            Case "三个月内"
                mdatBeginDate = CDate(Format(DateAdd("M", -3, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = Sys.Currentdate
            Case "半年内"
                mdatBeginDate = CDate(Format(DateAdd("M", -6, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = Sys.Currentdate
            Case "一年内"
                mdatBeginDate = CDate(Format(DateAdd("yyyy", -1, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = Sys.Currentdate
            Case "自定义"
                mdatBeginDate = CDate(Format(dtpBeginDate, "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = CDate(Format(dtpEndDate, "yyyy-mm-dd") & " 23:59:59")
        End Select
    End With
End Sub

Private Sub cbo库房_Click()
    mlng库房ID = cbo库房.ItemData(cbo库房.ListIndex)
    If cbo库房.Text = "所有库房" Then
        vsfLeft.ColHidden(vsfLeft.ColIndex("库房")) = False
    Else
        vsfLeft.ColHidden(vsfLeft.ColIndex("库房")) = True
    End If
End Sub

Private Sub chkALLVisible1_Click()
    If vsfAll.rows = 1 Then Exit Sub
    chkALLVisible2.Value = chkALLVisible1.Value
    Call SetVsfDta(1)
    Call SetDetailsData
End Sub

Private Sub chkALLVisible2_Click()
    chkALLVisible1.Value = chkALLVisible2.Value
    If vsfAll.rows = 1 Then Exit Sub
    Call SetVsfDta(1)
    Call SetDetailsData
End Sub

Private Sub cmdFind_Click()
    '提取数据代码
    Dim datBeginDate As Date
    Dim datEndDate As Date
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    
    vsfAll.rows = 1
    vsfList.rows = 1
    If cboDate.Text = "自定义" Then
        mdatBeginDate = CDate(Format(dtpBeginDate, "yyyy-mm-dd") & " 00:00:00")
        mdatEndDate = CDate(Format(dtpEndDate, "yyyy-mm-dd") & " 23:59:59")
    End If
    If ActiveControl Is cmdFind Then
        txtNo.Text = ""
        If cbo库房.Text = "所有库房" Then
            gstrSQL = ""
        Else
            gstrSQL = "  And A.库房id=[3]"
        End If
        gstrSQL = "Select b.名称, a.原始no, a.上次no, a.本次no As NO, a.审核人, a.审核日期" & vbNewLine & _
                "From 药品财务审核 A, 部门表 B" & vbNewLine & _
                "Where a.库房id = b.Id And a.单据 = 1 " & gstrSQL & " And a.审核日期 Between [1] And [2]" & vbNewLine & _
                "Order By a.审核日期 Desc"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "财务审核查询", mdatBeginDate, mdatEndDate, mlng库房ID)
    Else
        If cbo库房.Text = "所有库房" Then
            gstrSQL = ""
        Else
            gstrSQL = " And A.库房id=[2]"
        End If
        gstrSQL = "Select b.名称, a.原始no, a.上次no, a.本次no As NO, a.审核人, a.审核日期" & vbNewLine & _
                "From 药品财务审核 A, 部门表 B" & vbNewLine & _
                "Where a.库房id = b.Id And 单据 = 1" & gstrSQL & " And 本次no = [1]" & vbNewLine & _
                "Order By 审核日期 Desc"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "财务审核查询", txtNo.Text, mlng库房ID)
    End If
    
    vsfLeft.rows = 1
    If rsTemp.RecordCount > 0 Then
        rsTemp.Sort = " no asc"
        With vsfLeft
            .rows = rsTemp.RecordCount + 1
            For lngRow = 1 To rsTemp.RecordCount
                .TextMatrix(lngRow, .ColIndex("库房")) = rsTemp!名称
                .TextMatrix(lngRow, .ColIndex("原始NO")) = rsTemp!原始NO
                .TextMatrix(lngRow, .ColIndex("上次no")) = rsTemp!上次no
                .TextMatrix(lngRow, .ColIndex("no")) = rsTemp!NO
                .TextMatrix(lngRow, .ColIndex("审核人")) = rsTemp!审核人
                .TextMatrix(lngRow, .ColIndex("审核时间")) = Format(rsTemp!审核日期, "yyyy-mm-dd")
                rsTemp.MoveNext
            Next
        End With
    End If
End Sub

Private Sub GetALLData()
    '获取汇总信息
    Dim strsql As String
    Dim str原始NO As String
    
    On Error GoTo errHandle
    Set mrsData = Nothing
    If vsfLeft.rows = 1 Then Exit Sub
    gstrSQL = "Select '冲销单据' As 类型, a.No, a.药品id, c.名称, c.编码, c.规格, a.产地,a.审核日期, c.计算单位, d.门诊单位, d.门诊包装, d.住院单位, d.住院包装, d.药库单位, d.药库包装, a.批号," & vbNewLine & _
        "       a.实际数量, a.成本价, a.成本金额, a.零售价, a.零售金额, a.差价, e.发票号,e.发票代码,e.发票日期,e.发票金额,a.摘要" & vbNewLine & _
        "From 药品收发记录 A, 药品财务审核 B, 收费项目目录 C, 药品规格 D, 应付记录 E" & vbNewLine & _
        "Where a.No = b.本次no And a.药品id = c.Id And c.Id = d.药品id And a.Id = e.收发id(+) And a.单据 = 1 And b.原始no =[1] And" & vbNewLine & _
        "      a.审核日期 Is Not Null And (Mod(a.记录状态, 3) = 0 Or a.记录状态 =1)" & vbNewLine & _
        "Union" & vbNewLine & _
        "Select '原始单据' As 类型, a.No, a.药品id, c.名称, c.编码, c.规格, a.产地,a.审核日期, c.计算单位, d.门诊单位, d.门诊包装, d.住院单位, d.住院包装, d.药库单位, d.药库包装, a.批号," & vbNewLine & _
        "       a.实际数量, a.成本价, a.成本金额, a.零售价, a.零售金额, a.差价, e.发票号, e.发票代码, e.发票日期, e.发票金额,a.摘要" & vbNewLine & _
        "From 药品收发记录 A, 收费项目目录 C, 药品规格 D, 应付记录 E" & vbNewLine & _
        "Where a.药品id = c.Id And c.Id = d.药品id And a.Id = e.收发id(+) And a.单据 = 1 And a.No = [1] And a.审核日期 Is Not Null And" & vbNewLine & _
        "      Mod(a.记录状态, 3) = 0"
    
    Set mrsData = zlDataBase.OpenSQLRecord(gstrSQL, "查询所有数据", vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("原始no")))
    Set mrsCloneDta = mrsData.Clone  '克隆数据集
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDetailsData()
    '获取明细数据
    
End Sub

Private Sub Form_Load()
    Me.Height = 600 * 15
    Me.Width = 800 * 15
    Call SetControlLocation
    Call SetCBOValue
    dtpBeginDate.Value = DateAdd("d", -7, Sys.Currentdate)
    dtpEndDate.Value = Sys.Currentdate
    
    Call GetDrugDigit(mlng库房ID, "药品外购入库管理", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
End Sub

Private Sub SetCBOValue()
    Dim arrtemp As Variant
    Dim i As Integer
    Dim strIndex As String
    Dim strTemp As String
    '为日期下拉框赋值
    With cboDate
        .AddItem "一个月内"
        .AddItem "三个月内"
        .AddItem "半年内"
        .AddItem "一年内"
        .AddItem "自定义"
        .ListIndex = 0
    End With
    
    ReDim arrtemp(UBound(Split(mStr库房, "|"))) As String
    
    With cbo库房
        .Clear
        .AddItem "所有库房"
        .ItemData(.NewIndex) = "0"
        For i = 0 To UBound(arrtemp) - 1
            strIndex = ""
            strTemp = ""
            arrtemp(i) = Split(mStr库房, "|")(i)
            strIndex = Mid(arrtemp(i), 1, InStr(1, arrtemp(i), ",") - 1)
            strTemp = Mid(arrtemp(i), InStr(1, arrtemp(i), ",") + 1)
            .AddItem strTemp
            .ItemData(.NewIndex) = strIndex
        Next
        
        .ListIndex = Val(mstr当前库房) + 1
    End With
End Sub

Private Sub Form_Resize()
    Call SetControlLocation
    If sstInfo.Tab = 0 Then
        vsfList.Visible = False
        vsfAll.Visible = True
    Else
        vsfList.Visible = True
        vsfAll.Visible = False
    End If
End Sub

Private Sub optFloor_Click()
    vsfList.ColHidden(vsfList.ColIndex("no")) = True
'    vsfList.ColHidden(vsfList.ColIndex("原始")) = True
    If vsfList.rows = 1 Then Exit Sub
    Call SetDetailsData
End Sub

Private Sub optMedi_Click()
    vsfList.ColHidden(vsfList.ColIndex("no")) = False
'    vsfList.ColHidden(vsfList.ColIndex("原始")) = False
    If vsfList.rows = 1 Then Exit Sub
    Call SetDetailsData
End Sub


Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        If picLeft.Width + x < 1000 Then Exit Sub
        If sstInfo.Width - x < 2000 Then Exit Sub
        picLeft.Width = picLeft.Width + x
        picSplit.Left = picSplit.Left + x
        sstInfo.Width = sstInfo.Width - x
        sstInfo.Left = sstInfo.Left + x
        vsfLeft.Width = picLeft.ScaleWidth - 120
        vsfAll.Width = sstInfo.Width - 100
        vsfList.Width = sstInfo.Width - 100
    End If
End Sub

Private Sub sstInfo_Click(PreviousTab As Integer)
    If sstInfo.Tab = 0 Then
        vsfList.Visible = False
        vsfAll.Visible = True
    Else
        vsfList.Visible = True
        vsfAll.Visible = False
    End If
End Sub

Private Sub txtNO_GotFocus()
    With txtNo
        .SelStart = 0
        .SelLength = Len(txtNo.Text)
    End With
End Sub

Private Sub TxtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intNO As Integer, strNo As String
    
    If KeyCode = vbKeyReturn Then
        '提取数据代码
        intNO = 21
        If KeyCode = vbKeyReturn Then
            If Len(txtNo) < 8 And Len(txtNo) > 0 Then
                txtNo.Text = zlCommFun.GetFullNO(txtNo.Text, intNO, mlng库房ID)
            End If
            Call cmdFind_Click
        End If
    End If
End Sub

Private Sub vsfLeft_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsfLeft.rows = 1 Then Exit Sub
    If OldRow <> NewRow Then
        Call GetALLData '查询数据
        If mrsData.RecordCount > 0 Then
            mrsData.Sort = " no asc"
            Call SetVsfDta(0) '赋值
            Call SetDetailsData
        End If
    End If
End Sub

Private Sub SetVsfDta(ByVal intModel As Integer)
    '为汇总和明细控件赋值
    '参数 intModel 0-点击列表查询 1-条件改变查询
    Dim lngRow As Long
    Dim lngCol As Long
    Dim str上次NO As String
    Dim strNewNO As String
    Dim str原始NO As String
    Dim strNo As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim arrNo As Variant
    Dim dbl采购金额合计 As Double
    Dim dbl售价金额合计 As Double
    Dim dbl差价金额合计 As Double
    Dim dbl发票金额合计 As Double
    Dim str单位 As String
    Dim str换算系数 As String
    Dim strNOType As String
    Dim str摘要 As String
    
    With vsfAll
        .rows = 1
        str上次NO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("上次NO"))
        strNewNO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("NO"))
        str原始NO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("原始NO"))
        If intModel = 1 Then
            Set mrsData = Nothing
            Set mrsData = mrsCloneDta.Clone
            mrsData.Sort = " no asc"
        End If
        
        If chkALLVisible1.Value = 1 Then '显示完整单据
            '获取完整单据
            mrsData.MoveFirst
            Do While Not mrsData.EOF
                strTemp = mrsData!NO
                If InStr(1, "," & strNo & ",", "," & strTemp & ",") = 0 Then
                    strNo = strNo & "," & strTemp
                End If
                mrsData.MoveNext
            Loop
            If strNo <> "" Then
                strNo = Mid(strNo, 2)
                arrNo = Split(strNo, ",")
                For i = 0 To UBound(arrNo)
                    strTemp = " no='" & arrNo(i) & "'"
                    mrsData.Filter = strTemp
                    dbl采购金额合计 = 0
                    dbl售价金额合计 = 0
                    dbl差价金额合计 = 0
                    dbl发票金额合计 = 0
                    Do While Not mrsData.EOF
                        strNOType = mrsData!类型
                        str摘要 = IIf(IsNull(mrsData!摘要), "", mrsData!摘要)
                        dbl采购金额合计 = dbl采购金额合计 + mrsData!成本金额
                        dbl售价金额合计 = dbl售价金额合计 + mrsData!零售金额
                        dbl差价金额合计 = dbl差价金额合计 + mrsData!差价
                        dbl发票金额合计 = dbl发票金额合计 + IIf(IsNull(mrsData!发票金额), 0, mrsData!发票金额)
                        mrsData.MoveNext
                    Loop
                    .rows = .rows + 1
                    .Cell(flexcpPicture, .rows - 1, .ColIndex("原始"), .rows - 1, .ColIndex("原始")) = IIf(strNOType = "原始单据", imgPicture.ListImages(1).Picture, "")
                    .TextMatrix(.rows - 1, .ColIndex("摘要")) = str摘要
                    .TextMatrix(.rows - 1, .ColIndex("no")) = arrNo(i)
                    .TextMatrix(.rows - 1, .ColIndex("采购金额")) = zlStr.FormatEx(dbl采购金额合计, mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, .ColIndex("售价金额")) = zlStr.FormatEx(dbl售价金额合计, mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, .ColIndex("差价")) = zlStr.FormatEx(dbl差价金额合计, mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, .ColIndex("发票金额")) = zlStr.FormatEx(dbl发票金额合计, mintMoneyDigit, , True)
                Next
            End If
        Else '只显示当前单据和产生当前单据的冲销原始单据
            For i = 1 To 2
                If i = 1 Then
                    strTemp = " no='" & str上次NO & "'"
                Else
                    strTemp = " no='" & strNewNO & "'"
                End If
                dbl采购金额合计 = 0
                dbl售价金额合计 = 0
                dbl差价金额合计 = 0
                dbl发票金额合计 = 0
                mrsData.Filter = strTemp
                Do While Not mrsData.EOF
                    strNOType = mrsData!类型
                    str摘要 = IIf(IsNull(mrsData!摘要), "", mrsData!摘要)
                    dbl采购金额合计 = dbl采购金额合计 + mrsData!成本金额
                    dbl售价金额合计 = dbl售价金额合计 + mrsData!零售金额
                    dbl差价金额合计 = dbl差价金额合计 + mrsData!差价
                    dbl发票金额合计 = dbl发票金额合计 + IIf(IsNull(mrsData!发票金额), 0, mrsData!发票金额)
                    mrsData.MoveNext
                Loop
                .rows = .rows + 1
                .Cell(flexcpPicture, .rows - 1, .ColIndex("原始"), .rows - 1, .ColIndex("原始")) = IIf(strNOType = "原始单据", imgPicture.ListImages(1).Picture, "")
                .TextMatrix(.rows - 1, .ColIndex("摘要")) = str摘要
                .TextMatrix(.rows - 1, .ColIndex("no")) = IIf(i = 1, str上次NO, strNewNO)
                .TextMatrix(.rows - 1, .ColIndex("采购金额")) = zlStr.FormatEx(dbl采购金额合计, mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, .ColIndex("售价金额")) = zlStr.FormatEx(dbl售价金额合计, mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, .ColIndex("差价")) = zlStr.FormatEx(dbl差价金额合计, mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, .ColIndex("发票金额")) = zlStr.FormatEx(dbl发票金额合计, mintMoneyDigit, , True)
            Next
        End If
        Call CheckValue
    End With
End Sub

Private Sub SetDetailsData()
    '为明细表格赋值
    '为汇总和明细控件赋值
    '参数 intModel 0-点击列表查询 1-条件改变查询
    Dim lngRow As Long
    Dim lngCol As Long
    Dim str上次NO As String
    Dim strNewNO As String
    Dim str原始NO As String
    Dim strNo As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim arrNo As Variant
    Dim dbl采购金额合计 As Double
    Dim dbl售价金额合计 As Double
    Dim dbl差价金额合计 As Double
    Dim dbl发票金额合计 As Double
    Dim str单位 As String
    Dim str换算系数 As String
    Dim dbl发票金额 As Double
    Dim strNOType As String
    
    With vsfList
        .rows = 1
        str上次NO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("上次NO"))
        strNewNO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("NO"))
        str原始NO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("原始NO"))
        
        Set mrsData = mrsCloneDta.Clone
        
        If chkALLVisible1.Value = 1 Then '显示完整单据
        Else
            strTemp = " no='" & str上次NO & "' or no='" & strNewNO & " '"
            mrsData.Filter = strTemp
        End If
        '获取完整单据
        If optFloor.Value = True Then '按照单据分组
            mrsData.Sort = " no asc"
        Else
            mrsData.Sort = " 药品id,no asc"
        End If
        
        mrsData.MoveFirst
        Do While Not mrsData.EOF
            vsfList.rows = vsfList.rows + 1
            If optFloor.Value = True Then '按照单据分组
                If vsfList.rows > 2 Then
                    If mrsData!NO <> vsfList.TextMatrix(vsfList.rows - 2, vsfList.ColIndex("no")) Then
                        vsfList.MergeCells = flexMergeFree
                        vsfList.MergeRow(vsfList.rows - 1) = True
                        vsfList.Cell(flexcpText, vsfList.rows - 1, 0, vsfList.rows - 1, vsfList.Cols - 1) = "NO：" & vsfList.TextMatrix(vsfList.rows - 2, vsfList.ColIndex("no"))
                        vsfList.Cell(flexcpBackColor, vsfList.rows - 1, 0, vsfList.rows - 1, vsfList.Cols - 1) = &HFFFFFF  ' &HFFC0C0
                        vsfList.rows = vsfList.rows + 1
                    End If
                End If
            End If
            Select Case mintUnit
                Case M_INT_售价单位
                    str单位 = mrsData!计算单位
                    str换算系数 = 1
                Case M_INT_门诊单位
                    str单位 = mrsData!门诊单位
                    str换算系数 = mrsData!门诊包装
                Case M_INT_住院单位
                    str单位 = mrsData!住院单位
                    str换算系数 = mrsData!住院包装
                Case M_INT_药库单位
                    str单位 = mrsData!药库单位
                    str换算系数 = mrsData!药库包装
            End Select
            strNOType = mrsData!类型
            vsfList.Cell(flexcpPicture, vsfList.rows - 1, vsfList.ColIndex("原始"), vsfList.rows - 1, vsfList.ColIndex("原始")) = IIf(strNOType = "原始单据", imgPicture.ListImages(1).Picture, "")
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("no")) = mrsData!NO
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("药品id")) = mrsData!药品id
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("编码药品名称和规格")) = "[" & mrsData!编码 & "]" & mrsData!名称 & "(" & IIf(IsNull(mrsData!规格), "", mrsData!规格) & ")"
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("产地批号")) = IIf(IsNull(mrsData!产地), "", mrsData!产地) & "(" & IIf(IsNull(mrsData!批号), "", mrsData!批号) & ")"
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("数量")) = zlStr.FormatEx(mrsData!实际数量 / str换算系数, mintNumberDigit, , True) & "(" & str单位 & ")"
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("采购价")) = zlStr.FormatEx(mrsData!成本价 * str换算系数, mintCostDigit, , True)
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("采购金额")) = zlStr.FormatEx(mrsData!成本金额, mintMoneyDigit, , True)
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("售价")) = zlStr.FormatEx(mrsData!零售价 * str换算系数, mintPriceDigit, , True)
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("售价金额")) = zlStr.FormatEx(mrsData!零售金额, mintMoneyDigit, , True)
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("差价")) = zlStr.FormatEx(mrsData!差价, mintMoneyDigit, , True)
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("发票号")) = IIf(IsNull(mrsData!发票号), "", mrsData!发票号)
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("发票代码")) = IIf(IsNull(mrsData!发票代码), "", mrsData!发票代码)
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("发票日期")) = IIf(IsNull(mrsData!发票日期), "", Format(mrsData!发票日期, "yyyy-mm-dd"))
            dbl发票金额 = IIf(IsNull(mrsData!发票金额), 0, mrsData!发票金额)
            vsfList.TextMatrix(vsfList.rows - 1, .ColIndex("发票金额")) = IIf(dbl发票金额 = 0, "", zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True))
            
            mrsData.MoveNext
        Loop
        If optFloor.Value = True Then '按照单据分组
            vsfList.rows = vsfList.rows + 1
            vsfList.MergeCells = flexMergeFree
            vsfList.MergeRow(vsfList.rows - 1) = True
            vsfList.Cell(flexcpText, vsfList.rows - 1, 0, vsfList.rows - 1, vsfList.Cols - 1) = "NO：" & vsfList.TextMatrix(vsfList.rows - 2, vsfList.ColIndex("no"))
            vsfList.Cell(flexcpBackColor, vsfList.rows - 1, 0, vsfList.rows - 1, vsfList.Cols - 1) = &HFFFFFF  ' &HFFC0C0
        End If
        Call CheckValue
    End With
End Sub

Private Sub vsfLeft_EnterCell()
    With vsfLeft
        .FocusRect = flexFocusSolid
    End With
End Sub

Private Sub CheckValue()
    Dim lngRow As Long
    Dim i As Long
    Dim lngCol As Long
    '检查表格中哪些信息不相同，同列不相同内容用红色字体标注
    '汇总表格
    With vsfAll
        For lngRow = 2 To .rows - 1
            For lngCol = 2 To .Cols - 1
                If .TextMatrix(1, lngCol) <> .TextMatrix(lngRow, lngCol) Then
                    .Cell(flexcpForeColor, lngRow, lngCol, lngRow, lngCol) = vbRed
                End If
            Next
        Next
    End With
    '明细表格
    With vsfList
        If .rows < 3 Then Exit Sub
        
        For lngRow = 1 To .rows - 1
            For i = lngRow + 1 To .rows - 1
                If i > .rows - 1 Then Exit For
                If .TextMatrix(lngRow, .ColIndex("药品id")) = .TextMatrix(i, .ColIndex("药品id")) Then
                    For lngCol = 3 To .Cols - 1
                        If .TextMatrix(lngRow, lngCol) <> .TextMatrix(i, lngCol) Then
                            .Cell(flexcpForeColor, i, lngCol, i, lngCol) = vbRed
                        End If
                    Next
                End If
            Next
        Next
    End With
End Sub

Public Sub showMe(ByVal frmPar As Form, ByVal str库房 As String, ByVal str当前库房 As Long)
    mStr库房 = str库房
    mstr当前库房 = str当前库房
    Me.Show vbModal, frmPar
End Sub

