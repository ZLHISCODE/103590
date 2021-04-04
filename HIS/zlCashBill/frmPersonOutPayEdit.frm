VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmPersonOutPayEdit 
   BorderStyle     =   0  'None
   Caption         =   "借出管理"
   ClientHeight    =   7545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList ils16 
      Left            =   7170
      Top             =   1005
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":0000
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":059A
            Key             =   "OutPay"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":0B34
            Key             =   "Requisition"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":10CE
            Key             =   "Out"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils24 
      Left            =   5805
      Top             =   735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":1668
            Key             =   "Requisition"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":1D62
            Key             =   "OutPay"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":245C
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picOutPay 
      BorderStyle     =   0  'None
      Height          =   3330
      Left            =   4005
      ScaleHeight     =   3330
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   1620
      Width           =   3555
      Begin VSFlex8Ctl.VSFlexGrid vsOutPay 
         Height          =   2145
         Left            =   135
         TabIndex        =   2
         Top             =   585
         Width           =   2895
         _cx             =   5106
         _cy             =   3784
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPersonOutPayEdit.frx":2B56
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
         ExplorerBar     =   7
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
         VirtualData     =   -1  'True
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
         Begin VB.PictureBox picImgOutPay 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   30
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   4
            Top             =   60
            Width           =   210
            Begin VB.Image imgColSel 
               Height          =   195
               Left            =   0
               Picture         =   "frmPersonOutPayEdit.frx":2CAC
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin VB.PictureBox picLoanRequisition 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   75
      ScaleHeight     =   4695
      ScaleWidth      =   3450
      TabIndex        =   0
      Top             =   885
      Width           =   3450
      Begin VSFlex8Ctl.VSFlexGrid vsRequisition 
         Height          =   2145
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2895
         _cx             =   5106
         _cy             =   3784
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPersonOutPayEdit.frx":31FA
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
         ExplorerBar     =   7
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
         VirtualData     =   -1  'True
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
         Begin VB.PictureBox picImgRequisition 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   5
            Top             =   60
            Width           =   210
            Begin VB.Image imgColRequisition 
               Height          =   195
               Left            =   0
               Picture         =   "frmPersonOutPayEdit.frx":32E3
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPersonOutPayEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private mstrPrivs As String, mlngModule As Long
Private mArrFilter As Variant   '过滤条件
Private mcbsThis As Object
Private Const conPane_Requisition = 0
Private Const conPane_OutPay = 1

Public Function zlReLoadData(ByVal mcllFilter As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新加载数据
    '返回:加载成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2009-09-07 14:43:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set mArrFilter = mcllFilter
    Call LoadDataToRpt
    Call LoadRequisition
    zlReLoadData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2009-09-09 15:45:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsRequisition
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColHidden(.ColIndex("ID")) = True
        .ColData(.ColIndex("标志")) = "-1|1"
        .ColData(.ColIndex("借款人")) = "1|0"
        .ColData(.ColIndex("借款金额")) = "1|0"
        .ColData(.ColIndex("申请时间")) = "1|0"
    End With
    With vsOutPay
        .ColHidden(.ColIndex("ID")) = True
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("标志")) = "-1|1"
        .ColData(.ColIndex("借款人")) = "1|0"
        .ColData(.ColIndex("借款金额")) = "1|0"
        .ColData(.ColIndex("申请时间")) = "1|0"
    End With
    
End Sub
Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, lngRow As Long
    Dim blnHistory As Boolean, strStartDate As String
    
    
    Err = 0: On Error GoTo ErrHand:
    zlCommFun.ShowFlash "正在装载借款数据,请稍后..."
    strStartDate = "3000-01-01 00:00:00"
    If strStartDate > CStr(mArrFilter("借出-申请时间")(0)) And CStr(mArrFilter("借出-申请时间")(0)) <> "1901-01-01" Then
        strStartDate = CStr(mArrFilter("借出-申请时间")(0))
    End If
    If strStartDate > CStr(mArrFilter("借出-借出时间")(0)) And CStr(mArrFilter("借出-借出时间")(0)) <> "1901-01-01" Then
        strStartDate = CStr(mArrFilter("借出-借出时间")(0))
    End If
    If strStartDate > CStr(mArrFilter("借出-取消时间")(0)) And CStr(mArrFilter("借出-取消时间")(0)) <> "1901-01-01" Then
        strStartDate = CStr(mArrFilter("借出-取消时间")(0))
    End If
    
    If strStartDate <> "3000-01-01 00:00:00" Then blnHistory = zlDatabase.DateMoved(strStartDate, , , Me.Caption)


    If CStr(mArrFilter("借出-申请时间")(0)) <> "1901-01-01" And CStr(mArrFilter("借出-借出时间")(0)) <> "1901-01-01" And CStr(mArrFilter("借出-取消时间")(0)) <> "1901-01-01" Then
        strFilter = "   (申请时间 between [1] and [2] or 借出时间 between [3] and [4] or 取消时间 between [5] and [6]) and 借出时间 is not Null   "
         
    ElseIf CStr(mArrFilter("借出-申请时间")(0)) <> "1901-01-01" And CStr(mArrFilter("借出-借出时间")(0)) <> "1901-01-01" And CStr(mArrFilter("借出-取消时间")(0)) = "1901-01-01" Then
        strFilter = "   (申请时间 between [1] and [2] or 借出时间 between [3] and [4] and 借出时间 is not Null )   "
    ElseIf CStr(mArrFilter("借出-申请时间")(0)) <> "1901-01-01" And CStr(mArrFilter("借出-借出时间")(0)) = "1901-01-01" And CStr(mArrFilter("借出-取消时间")(0)) <> "1901-01-01" Then
        strFilter = "   (申请时间 between [1] and [2] or 取消时间 between [5] and [6]) and 借出时间 is not Null   "
    ElseIf CStr(mArrFilter("借出-申请时间")(0)) = "1901-01-01" And CStr(mArrFilter("借出-借出时间")(0)) <> "1901-01-01" And CStr(mArrFilter("借出-取消时间")(0)) <> "1901-01-01" Then
        strFilter = "   ( 借出时间 between [3] and [4] or 取消时间 between [5] and [6]) and 借出时间 is not Null   "
    ElseIf CStr(mArrFilter("借出-申请时间")(0)) <> "1901-01-01" Then
        strFilter = "   (申请时间 between [1] and [2]   ) and 借出时间 is not Null "
    ElseIf CStr(mArrFilter("借出-借出时间")(0)) <> "1901-01-01" Then
        strFilter = "   (借出时间 between [3] and [4])"
    Else
        strFilter = "   (取消时间 between [5] and [6] )"
    End If
    
    If CStr(mArrFilter("借款人")) <> "" Then strFilter = strFilter & " and 借款人 like [7]"
    strFilter = strFilter & " and 借出人 like [8]"
    
    gstrSQL = " " & _
    "    Select distinct  A.Id, A.借款金额, A.备注, A.借款人, to_char(A.申请时间,'yyyy-mm-dd hh24:mi:ss') as 申请时间 ,  " & _
    "           A.借出人, to_char(A.借出时间,'yyyy-mm-dd hh24:mi:ss') as 借出时间, " & _
    "           to_char(A.取消时间,'yyyy-mm-dd hh24:mi:ss') as 取消时间, A.取消原因,Decode(B.记录ID,NULL,0,1) as 已缴款" & _
    "    From 人员借款记录 A,人员收缴对照 B " & _
    "    Where A.ID=B.记录ID(+) and B.性质(+)=4 And " & strFilter
    If blnHistory Then
        gstrSQL = gstrSQL & vbCrLf & " Union ALL " & Replace(Replace(gstrSQL, "人员借款记录", "H人员借款记录"), "Decode(B.记录ID,NULL,0,1)", " 2 ") & vbCrLf
    End If
    gstrSQL = gstrSQL & _
    "    Order by   借出时间, 借出人 "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        CDate(mArrFilter("借出-申请时间")(0)), CDate(mArrFilter("借出-申请时间")(1)), _
        CDate(mArrFilter("借出-借出时间")(0)), CDate(mArrFilter("借出-借出时间")(1)), _
        CDate(mArrFilter("借出-取消时间")(0)), CDate(mArrFilter("借出-取消时间")(1)), _
        CStr(mArrFilter("借款人")), UserInfo.姓名)
    
    With Me.vsOutPay
        .Clear 1
        .Rows = 2: lngRow = 1
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("借款人")) = Nvl(rsTemp!借款人)
            .TextMatrix(lngRow, .ColIndex("借款金额")) = Format(Val(Nvl(rsTemp!借款金额)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("申请时间")) = Nvl(rsTemp!申请时间)
            .TextMatrix(lngRow, .ColIndex("备注")) = Nvl(rsTemp!备注)
            .TextMatrix(lngRow, .ColIndex("借出人")) = Nvl(rsTemp!借出人)
            .TextMatrix(lngRow, .ColIndex("借出时间")) = Nvl(rsTemp!借出时间)
            .TextMatrix(lngRow, .ColIndex("取消时间")) = Nvl(rsTemp!取消时间)
            .TextMatrix(lngRow, .ColIndex("取消原因")) = Nvl(rsTemp!取消原因)
            If Nvl(rsTemp!取消时间) <> "" Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
                .Cell(flexcpPicture, lngRow, .ColIndex("借款人")) = ils16.ListImages("Cancel").Picture
            Else
                .Cell(flexcpPicture, lngRow, .ColIndex("借款人")) = ils16.ListImages("OutPay").Picture
            End If
            If Val(Nvl(rsTemp!已缴款)) = 1 Then
                .Cell(flexcpPicture, lngRow, .ColIndex("标志")) = ils16.ListImages("Out").Picture
            ElseIf Val(Nvl(rsTemp!已缴款)) = 2 Then
                '已经转成历史数据
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H8000000F
            End If
            .Cell(flexcpData, lngRow, .ColIndex("标志")) = Val(Nvl(rsTemp!已缴款))
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsOutPay, "借款管理", "借出列表", True
        .ColWidth(.ColIndex("标志")) = 285
    End With
    zlCommFun.StopFlash
   Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
     Me.vsOutPay.Redraw = flexRDBuffered
End Sub
Private Sub LoadRequisition()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载请求数据
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, lngRow As Long
    
    Err = 0: On Error GoTo ErrHand:
    
    zlCommFun.ShowFlash "正在装载借款申请,请稍后..."
    
    gstrSQL = " " & _
    "    Select Id, 借款金额, 备注, 借款人, to_char(申请时间,'yyyy-mm-dd hh24:mi:ss') as 申请时间 ,  " & _
    "           借出人, to_char(借出时间,'yyyy-mm-dd hh24:mi:ss') as 借出时间, " & _
    "           to_char(取消时间,'yyyy-mm-dd hh24:mi:ss') as 取消时间, 取消原因" & _
    "    From 人员借款记录 " & _
    "    Where 借出时间 is null and 借出人=[1] " & _
    "    Order by  申请时间,借款人 "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.姓名)
    
    With Me.vsRequisition
        .Clear 1
        .Rows = 2: .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("借款人")) = Nvl(rsTemp!借款人)
            .TextMatrix(lngRow, .ColIndex("借款金额")) = Format(Val(Nvl(rsTemp!借款金额)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("申请时间")) = Nvl(rsTemp!申请时间)
            .TextMatrix(lngRow, .ColIndex("备注")) = Nvl(rsTemp!备注)
            .TextMatrix(lngRow, .ColIndex("借出人")) = Nvl(rsTemp!借出人)
            .Cell(flexcpPicture, lngRow, .ColIndex("借款人")) = ils16.ListImages("Requisition").Picture
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '恢复列设置
         zl_vsGrid_Para_Restore mlngModule, vsRequisition, "借款管理", "申请列表", True
        .ColWidth(.ColIndex("标志")) = 285
        .Redraw = flexRDBuffered
    End With
    zlCommFun.StopFlash
   Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
     Me.vsRequisition.Redraw = flexRDBuffered
End Sub
Private Sub InitPancel()
    Dim sngWidth As Single
    Dim strReg As String
    Dim panThis As Pane
    
    Set panThis = dkpMan.CreatePane(conPane_Requisition, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "借款申请信息"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Tag = picLoanRequisition
    
    Set panThis = dkpMan.CreatePane(conPane_OutPay, 250, 580, DockRightOf, Nothing)
    panThis.Title = "借出信息"
    panThis.Tag = conPane_OutPay
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Sub
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Requisition
        Item.Handle = picLoanRequisition.hwnd
    Case conPane_OutPay
        Item.Handle = Me.picOutPay.hwnd
    End Select
End Sub
Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    
    Call InitPancel
    Call InitVsGrid
    Call vsOutPay_LostFocus: Call vsRequisition_LostFocus
    vsRequisition_GotFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    zl_vsGrid_Para_Save mlngModule, vsOutPay, "借款管理", "借出列表", True
    zl_vsGrid_Para_Save mlngModule, vsRequisition, "借款管理", "申请列表", True
End Sub

Private Sub imgColRequisition_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlcontrol.GetControlRect(picImgRequisition.hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgRequisition.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRequisition, lngLeft, lngTop, imgColSel.Height)
    zl_vsGrid_Para_Save mlngModule, vsRequisition, Me.Name, "申请列表", True
End Sub

Private Sub imgColSel_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlcontrol.GetControlRect(picImgOutPay.hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgOutPay.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsOutPay, lngLeft, lngTop, imgColSel.Height)
    zl_vsGrid_Para_Save mlngModule, vsOutPay, Me.Name, "借出列表", True
End Sub

Private Sub picImgOutPay_Click()
    Call imgColSel_Click
End Sub

Private Sub picLoanRequisition_Resize()
    Err = 0: On Error Resume Next
    With picLoanRequisition
        vsRequisition.Left = .ScaleLeft
        vsRequisition.Width = .ScaleWidth
        vsRequisition.Top = .ScaleTop
        vsRequisition.Height = .ScaleHeight
    End With
End Sub
Private Sub picOutPay_Resize()
    Err = 0: On Error Resume Next
    With picOutPay
        vsOutPay.Left = .ScaleLeft
        vsOutPay.Width = .ScaleWidth
        vsOutPay.Top = .ScaleTop
        vsOutPay.Height = .ScaleHeight
    End With
End Sub


Public Function zlDefCommandBars(ByVal cbsThis As Object) As Boolean
    '----------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/1/9
    '----------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    Set mcbsThis = cbsThis
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_BillPrintSet, "借款单打印设置"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "确认借出(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeOff, "取消借出(&M)")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_Audit
        .Add FCONTROL, Asc("M"), conMenu_Edit_ChargeOff
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "确认借出"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeOff, "取消借出")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
     zlcontrol.ControlSetFocus vsRequisition
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ExcuteFunction(Optional ByVal blnOutPay As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:确认借出
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-09 12:04:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    
    If blnOutPay Then
        If zlStr.IsHavePrivs(mstrPrivs, "借款确认") = False Then Exit Sub
        With vsRequisition
            If .Row < 0 Or .Row > .Rows - 1 Then Exit Sub
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End With
    Else
        If zlStr.IsHavePrivs(mstrPrivs, "取消借出") = False Then Exit Sub
        With vsOutPay
            If .Row < 0 Or .Row > .Rows - 1 Then Exit Sub
            If Trim(.TextMatrix(.Row, .ColIndex("取消时间"))) <> "" Then Exit Sub
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End With
    End If
    If lngID = 0 Then Exit Sub
    If frmPersonLoanRequisitionEdit.ShowEdit(Me, IIf(blnOutPay, FN_借出, FN_取消借出), mstrPrivs, mlngModule, lngID) = False Then Exit Sub
    '重新刷新数据
    Call zlReLoadData(mArrFilter)
End Sub
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long
    Dim lngID  As Long
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_Edit_Audit   '确认借出
        Call ExcuteFunction(True)
    Case conMenu_Edit_ChargeOff    '取消借出
        Call ExcuteFunction(False)
    Case conMenu_View_Refresh   '刷新
        '重新刷新数据
        Call zlReLoadData(mArrFilter)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            If Me.ActiveControl Is vsRequisition Then
                With vsOutPay
                        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
                End With
            Else
                With vsOutPay
                        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
                End With
            End If
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "ID=" & lngID)
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Function HaveData() As Boolean
    '功能:是否有数据
    If Me.ActiveControl Is vsRequisition Then
        With Me.vsRequisition
            HaveData = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
        End With
    Else
        With Me.vsOutPay
            HaveData = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
        End With
    End If
End Function

Private Function GetOutPayStaut() As Boolean
    '功能:借出状态
    If Me.ActiveControl Is vsRequisition Then
        With Me.vsRequisition
            GetOutPayStaut = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
        End With
    Else
        With Me.vsOutPay
            GetOutPayStaut = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
        End With
    End If
End Function

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long
    If Me.Visible = False Then Exit Sub
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = HaveData
    Case conMenu_Edit_Audit '借出确认
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "借款确认")
            Control.Enabled = Control.Visible And Val(vsRequisition.TextMatrix(vsRequisition.Row, vsRequisition.ColIndex("ID"))) <> 0 And Me.ActiveControl Is vsRequisition
    Case conMenu_Edit_ChargeOff
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "取消借出")
        With Me.vsOutPay
             Control.Enabled = Control.Visible And Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0 And Me.ActiveControl Is vsOutPay
             Control.Enabled = Control.Enabled And Trim(.TextMatrix(.Row, .ColIndex("取消时间"))) = "" And Val(.Cell(flexcpData, .Row, .ColIndex("标志"))) = 0
        End With
    Case conMenu_View_Refresh
    End Select
End Sub
Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2009-09-09 11:24:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsGrid As VSFlexGrid
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr单位名称 & IIf(Not Me.ActiveControl Is vsRequisition, "借款清单", "借款申请清单")
    If Not Me.ActiveControl Is vsRequisition Then
        If CStr(mArrFilter("借出-申请时间")(0)) <> "1901-01-01" Then
            objRow.Add "申请时间：" & CStr(mArrFilter("借出-申请时间")(0)) & "至" & CStr(mArrFilter("借出-申请时间")(1))
        End If
        If CStr(mArrFilter("借出-借出时间")(0)) <> "1901-01-01" Then
            objRow.Add "借出时间：" & CStr(mArrFilter("借出-借出时间")(0)) & "至" & CStr(mArrFilter("借出-借出时间")(1))
        End If
        If CStr(mArrFilter("借出-取消时间")(0)) <> "1901-01-01" Then
            objRow.Add "取消时间：" & CStr(mArrFilter("借出-取消时间")(0)) & "至" & CStr(mArrFilter("借出-取消时间")(1))
        End If
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add "借出人：" & UserInfo.姓名
        If CStr(mArrFilter("借款人")) <> "" Then objRow.Add "借款人：" & mArrFilter("借出-借款人")
        objPrint.UnderAppRows.Add objRow
        Set vsGrid = vsOutPay
    Else
        objRow.Add "借出人：" & UserInfo.姓名
        objPrint.UnderAppRows.Add objRow
        Set vsGrid = vsRequisition
    End If
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("标志") Then .ColWidth(intCol) = 0
        Next
    End With
    
    Set objPrint.Body = IIf(Not Me.ActiveControl Is vsRequisition, vsOutPay, vsRequisition)
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub vsOutPay_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsOutPay
        If Col = .ColIndex("标志") Then Cancel = True
    End With
End Sub

Private Sub vsOutPay_DblClick()
        ExcuteFunction False
End Sub

Private Sub vsOutPay_DragDrop(Source As Control, x As Single, y As Single)
    If Source Is vsRequisition Then
        '拖动
        Call ExcuteFunction(True)   '确认
    End If
End Sub

Private Sub vsOutPay_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Static objIcon As IPictureDisp
    If Not Source Is vsOutPay Then
        If State = 0 Then
            Set objIcon = Source.DragIcon
        ElseIf State = 2 Then
            Set Source.DragIcon = ils16.ListImages("OutPay").Picture
        ElseIf State = 1 Then
            Set Source.DragIcon = objIcon
        End If
    End If
End Sub

Private Sub vsOutPay_GotFocus()
    vsOutPay.BackColorSel = &H8000000D
End Sub

Private Sub vsOutPay_LostFocus()
    vsOutPay.BackColorSel = &H8000000A
End Sub

Private Sub vsOutPay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    If Button <> 2 Then Exit Sub
    zlcontrol.ControlSetFocus vsOutPay, True
    Set objPopup = mcbsThis.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
End Sub

Private Sub vsOutPay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If zlStr.IsHavePrivs(mstrPrivs, "取消借出") = False Then Exit Sub
        If Val(vsOutPay.TextMatrix(vsOutPay.Row, vsOutPay.ColIndex("ID"))) = 0 Then Exit Sub
        If Trim(vsOutPay.TextMatrix(vsOutPay.Row, vsOutPay.ColIndex("取消时间"))) <> "" Then Exit Sub
        
        Set vsOutPay.DragIcon = ils16.ListImages("OutPay").Picture
        vsOutPay.Drag 1
    End If
End Sub
 

Private Sub vsRequisition_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRequisition
        If Col = .ColIndex("标志") Then Cancel = True
    End With
End Sub

Private Sub vsRequisition_DblClick()
    ExcuteFunction True
End Sub

Private Sub vsRequisition_DragDrop(Source As Control, x As Single, y As Single)
    If Source Is vsOutPay Then
        '拖动
        Call ExcuteFunction(False)    '确认
    End If
End Sub

Private Sub vsRequisition_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Static objIcon As IPictureDisp
    If Not Source Is vsRequisition Then
        If State = 0 Then
            Set objIcon = Source.DragIcon
        ElseIf State = 2 Then
            Set Source.DragIcon = ils16.ListImages("Requisition").Picture
        ElseIf State = 1 Then
            Set Source.DragIcon = objIcon
        End If
    End If
End Sub

Private Sub vsRequisition_GotFocus()
    vsRequisition.BackColorSel = &H8000000D
End Sub

Private Sub vsRequisition_LostFocus()
    vsRequisition.BackColorSel = &H8000000A
End Sub

Private Sub vsRequisition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    If Button <> 2 Then Exit Sub
    zlcontrol.ControlSetFocus vsRequisition, True
    Set objPopup = mcbsThis.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
End Sub

Private Sub vsRequisition_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If zlStr.IsHavePrivs(mstrPrivs, "借款确认") = False Then Exit Sub
        If Val(vsRequisition.TextMatrix(vsRequisition.Row, vsRequisition.ColIndex("ID"))) = 0 Then Exit Sub
        Set vsRequisition.DragIcon = ils16.ListImages("Requisition").Picture
        vsRequisition.Drag 1
    End If
End Sub

