VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDataMoveQuery 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   Icon            =   "frmDataMoveQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8745
   Begin MSComDlg.CommonDialog cdgSave 
      Left            =   810
      Top             =   1335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   30
      TabIndex        =   8
      Top             =   5580
      Width           =   8685
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   7185
         TabIndex        =   5
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&H)"
         Height          =   350
         Left            =   3900
         TabIndex        =   3
         ToolTipText     =   "查找：F3"
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   2535
         TabIndex        =   2
         Top             =   150
         Width           =   1320
      End
      Begin VB.ComboBox cboFind 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   1230
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "输出到&Excel"
         Height          =   350
         Left            =   5895
         TabIndex        =   4
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查找项目"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内容"
         Height          =   180
         Left            =   2145
         TabIndex        =   9
         Top             =   210
         Width           =   360
      End
   End
   Begin VB.Frame fraNote 
      Height          =   645
      Left            =   15
      TabIndex        =   6
      Top             =   -45
      Width           =   8700
      Begin VB.Label lblNote 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "###"
         Height          =   360
         Left            =   195
         TabIndex        =   7
         Top             =   180
         Width           =   8400
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsData 
      Height          =   4935
      Left            =   15
      TabIndex        =   0
      Top             =   630
      Width           =   8700
      _cx             =   15346
      _cy             =   8705
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
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   0
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   240
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   End
End
Attribute VB_Name = "frmDataMoveQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private mrsData As ADODB.Recordset
Private mintType As Integer
Private mdatBegin As Date
Private mdatEnd As Date
Private mstrTitle As String
Private mstrNote As String

Private mblnExcel As Boolean
Private mlngBegin As Long

Public Sub ShowMe(ByVal intType As Integer, ByVal datBegin As Date, ByVal datEnd As Date, ByVal strTitle As String, ByVal strNote As String, FrmParent As Object)
    mintType = intType
    mdatBegin = datBegin
    mdatEnd = datEnd
    mstrTitle = strTitle
    mstrNote = strNote
    
    On Error Resume Next
    Me.Show 1, FrmParent
End Sub

Private Sub cboFind_Click()
    mlngBegin = 0
    txtFind.Text = ""
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function HaveExcel() As Boolean
'功能：判断系统是否安装了Excel
'说明：同时初始化Excel对象
    Dim objExcel As Object
    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    HaveExcel = Err.Number = 0
    Set objExcel = Nothing
End Function

Private Sub cmdExcel_Click()
    Dim strFile As String
    Dim lngBack As Long, lngFore As Long
    
    strFile = Me.Caption & "(" & Format(mdatBegin, "yyyyMMdd") & "-" & Format(mdatEnd, "yyyyMMdd") & ").xls"
    On Error GoTo errH
    cdgSave.DialogTitle = "保存Excel表格"
    cdgSave.Filter = "Microsoft Office Excel文件(*.xls)|*.xls"
    cdgSave.flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
    cdgSave.FileName = strFile
    cdgSave.CancelError = True
    cdgSave.ShowSave
    On Error GoTo 0
    strFile = cdgSave.FileName
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName, "ExcelPath", Left(strFile, Len(strFile) - Len(cdgSave.FileTitle))
    
    vsData.redraw = flexRDNone
    lngBack = vsData.BackColorSel
    lngFore = vsData.ForeColorSel
    vsData.BackColorSel = vsData.BackColor
    vsData.ForeColorSel = vsData.ForeColor
    vsData.SaveGrid strFile, flexFileExcel, flexXLSaveFixedCells
    vsData.BackColorSel = lngBack
    vsData.ForeColorSel = lngFore
    vsData.redraw = flexRDDirect
    
    If mblnExcel Then
        ShellExecute Me.hwnd, "open", strFile, "", "", SW_SHOWNORMAL
    Else
        MsgBox "已经输出到文件""" & strFile & """中。", vbInformation, gstrSysName
    End If
errH:
End Sub

Private Sub cmdFind_Click()
    Dim lngRow As Long, blnFull As Boolean
    
    If txtFind.Text = "" Then
        MsgBox "请输入要查找的内容。", vbInformation, gstrSysName
        txtFind.SetFocus: Exit Sub
    End If
    
    lngRow = vsData.FindRow(txtFind.Text, mlngBegin + 1, cboFind.ItemData(cboFind.ListIndex), False, InStr("姓名,单据号,挂号单", cboFind.Text) = 0)
    If lngRow <> -1 Then
        mlngBegin = lngRow
        vsData.Row = lngRow
        Call vsData.ShowCell(vsData.Row, 0)
    Else
        mlngBegin = 0
        MsgBox "已经找到表格尾部，未发现符合条件的行。下次将重新从表头开始查找。", vbInformation, gstrSysName
    End If
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If cmdFind.Enabled Then cmdFind_Click
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName, mintType)
    Me.Caption = mstrTitle
    
    lblNote.Caption = "查询时间：" & Format(mdatBegin, "yyyy-MM-dd") & " 至 " & Format(mdatEnd, "yyyy-MM-dd")
    lblNote.Caption = lblNote.Caption & vbCrLf & mstrNote
    
    If Not LoadData Then Unload Me: Exit Sub
    mblnExcel = HaveExcel
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraNote.Left = 0
    fraNote.Top = -45
    fraNote.Width = Me.ScaleWidth
    lblNote.Width = fraNote.Width - lblNote.Left * 2
    
    vsData.Left = 0
    vsData.Top = fraNote.Top + fraNote.Height
    vsData.Width = Me.ScaleWidth
    vsData.Height = Me.ScaleHeight - vsData.Top - fraFunc.Height
    
    fraFunc.Left = 0
    fraFunc.Top = vsData.Top + vsData.Height
    fraFunc.Width = Me.ScaleWidth
    
    If fraFunc.Width - cmdCancel.Width - 500 >= 6500 Then
        cmdCancel.Left = fraFunc.Width - cmdCancel.Width - 500
    Else
        cmdCancel.Left = 6500
    End If
    cmdExcel.Left = cmdCancel.Left - cmdExcel.Width
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsData.State = 1 Then mrsData.Close
    Set mrsData = Nothing
    
    Call SaveWinState(Me, App.ProductName, mintType)
End Sub

Private Function LoadData() As Boolean
    Dim strBegin As String, strEnd As String
    Dim i As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    Set mrsData = New ADODB.Recordset
    
'    strBegin = "To_Date('" & Format(mdatBegin, "yyyy-MM-dd") & "','YYYY-MM-DD')"
'    strEnd = "To_Date('" & Format(mdatEnd, "yyyy-MM-dd") & "','YYYY-MM-DD')+1"
    strBegin = Format(mdatBegin, "yyyy-MM-dd")
    strEnd = Format(DateAdd("d", 1, mdatEnd), "yyyy-MM-dd")
    Select Case mintType
    Case 0
        '收费，挂号单据
        gstrSQL = _
        " Select Distinct  '门诊收费' As 单据类型, Decode(d.收费类别,'4','卫材未在此之前发料','药品未在此之前发药') as 无法转移原因," & _
        "       d.No As 单据号,d.标识号,d.姓名,d.性别,d.年龄,To_Char(d.登记时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间" & _
        " From 药品收发记录 l," & _
        "     ( Select d.Id,d.结帐id,d.NO,d.标识号,d.姓名,d.性别,d.年龄,d.登记时间,d.收费类别" & _
        "       From 门诊费用记录 d " & _
        "       Where d.登记时间>=[1] And d.登记时间<[2] And d.结帐ID Is Not Null" & _
        "             And d.记录性质 = 1 And d.收费类别 In ('4', '5', '6', '7')) d" & _
        " Where l.No = d.No And l.费用id = d.Id And Nvl(发药方式, 0) <> -1" & _
        "       And (l.审核日期 >=[2] Or l.审核日期 Is Null) And l.单据 In (8, 24)"
                          
        gstrSQL = gstrSQL & " Union ALL " & _
        " Select Distinct Decode(c.记录性质,1,'门诊收费',4,'门诊挂号') As 单据类型,'结算时使用的预交款未用完' as 无法转移原因," & _
        "       c.No As 单据号,c.标识号,c.姓名,c.性别,c.年龄,To_Char(c.登记时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间" & _
        " From 门诊费用记录 c,病人预交记录 d," & _
        "      (    Select d.No " & _
        "           From 病人预交记录 d," & _
        "               (Select 结帐id From 门诊费用记录  Where 登记时间>=[1] And 登记时间<[2] And 记录性质 In (1, 4) And Nvl(记帐费用,0)=0 ) l" & _
        "           Where d.结帐id = l.结帐id And d.记录性质 In (1, 11)" & _
        "           Group By d.No" & _
        "           Having d.No Is Not Null And Sum(d.金额) - Sum(d.冲预交) <> 0) n" & _
        " Where d.No = n.No And d.记录性质 In (1, 11)" & _
        " And c.结帐ID=d.结帐ID And c.记录性质 IN(1, 4) And Nvl(c.记帐费用,0)=0" & _
        " Order By 单据类型,单据号 Desc"
    Case 1
        '记帐单据
        gstrSQL = _
        " Select Distinct Decode(d.门诊标志,2,'住院记帐','门诊记帐') As 单据类型, Decode(d.收费类别,'4','卫材未在此之前发料','药品未在此之前发药') as 无法转移原因," & _
        "       d.No As 单据号,d.标识号,d.姓名,d.性别,d.年龄,To_Char(d.登记时间,'YYYY-MM-DD HH24:MI:SS') as 记帐时间" & _
        " From 药品收发记录 l," & _
        "     ( Select d.Id,d.结帐id,d.NO,d.标识号,d.姓名,d.性别,d.年龄,d.登记时间,d.收费类别,d.门诊标志" & _
        "       From 住院费用记录 d" & _
        "       Where d.登记时间>=[1] And 登记时间<[2] And d.结帐ID Is Not Null" & _
        "             And d.记录性质=2 And d.收费类别 In ('4', '5', '6', '7') " & _
        "       Union ALL " & _
        "       Select d.Id,d.结帐id,d.NO,d.标识号,d.姓名,d.性别,d.年龄,d.登记时间,d.收费类别,d.门诊标志" & _
        "       From 门诊费用记录 d" & _
        "       Where d.登记时间>=[1] And 登记时间<[2] And d.结帐ID Is Not Null" & _
        "             And d.记录性质=2 And d.收费类别 In ('4', '5', '6', '7') " & _
        "       ) d" & _
        " Where l.No = d.No And l.费用id = d.Id And Nvl(发药方式, 0) <> -1" & _
        "       And (l.审核日期 >= [2] Or l.审核日期 Is Null)  And l.单据 In (9, 10, 25, 26)"
        
        gstrSQL = gstrSQL & " Union ALL " & _
        " Select Distinct Decode(n.记录性质,2,Decode(d.门诊标志,2,'住院记帐','门诊记帐'),3,'自动记帐',5,'就诊卡记帐') As 单据类型,'同时结算的单据有未结清费用' as 无法转移原因," & _
        "       d.No As 单据号,d.标识号,d.姓名,d.性别,d.年龄,To_Char(d.登记时间,'YYYY-MM-DD HH24:MI:SS') as 记帐时间" & _
        " From 门诊费用记录 d," & _
        "     (   Select d.No, d.序号, Decode(d.记录性质, 12, 2, 13, 3, 15, 5, d.记录性质) As 记录性质" & _
        "         From 门诊费用记录 d, 病人结帐记录 l" & _
        "         Where d.结帐id = l.Id And l.收费时间>=[1] And l.收费时间<[2] " & _
        "               And d.记录性质 In (2, 12, 3, 13, 5, 15) And d.记帐费用 = 1" & _
        "         Group By d.No, d.序号, Decode(d.记录性质, 12, 2, 13, 3, 15, 5, d.记录性质)" & _
        "         Having d.No Is Not Null And d.序号 Is Not Null And Nvl(Sum(d.实收金额),0) - Nvl(Sum(d.结帐金额),0) <> 0 " & _
        "       ) n" & _
        " Where d.No = n.No And d.序号 = n.序号 And Decode(d.记录性质, 12, 2, 13, 3, 15, 5, d.记录性质) = n.记录性质"
        
        gstrSQL = gstrSQL & " Union ALL " & _
        " Select Distinct Decode(n.记录性质,2,Decode(d.门诊标志,2,'住院记帐','门诊记帐'),3,'自动记帐',5,'就诊卡记帐') As 单据类型,'同时结算的单据有未结清费用' as 无法转移原因," & _
        "       d.No As 单据号,d.标识号,d.姓名,d.性别,d.年龄,To_Char(d.登记时间,'YYYY-MM-DD HH24:MI:SS') as 记帐时间" & _
        " From 住院费用记录 d," & _
        "     (   Select d.No, d.序号, Decode(d.记录性质, 12, 2, 13, 3, 15, 5, d.记录性质) As 记录性质" & _
        "         From 住院费用记录 d, 病人结帐记录 l" & _
        "         Where d.结帐id = l.Id And l.收费时间>=[1] And l.收费时间<[2] " & _
        "               And d.记录性质 In (2, 12, 3, 13, 5, 15) And d.记帐费用 = 1" & _
        "           Group By d.No, d.序号, Decode(d.记录性质, 12, 2, 13, 3, 15, 5, d.记录性质)" & _
        "           Having d.No Is Not Null And d.序号 Is Not Null And Nvl(Sum(d.实收金额),0) - Nvl(Sum(d.结帐金额),0) <> 0 " & _
        "       ) n" & _
        " Where d.No = n.No And d.序号 = n.序号 And Decode(d.记录性质, 12, 2, 13, 3, 15, 5, d.记录性质) = n.记录性质"
        
        '理解有问题
        gstrSQL = gstrSQL & " Union ALL " & _
        " Select Distinct Decode(Mod(c.记录性质,10),2,Decode(c.门诊标志,2,'住院记帐','门诊记帐'),3,'自动记帐',5,'就诊卡记帐') As 单据类型, '结算时使用的预交款未用完' as 无法转移原因," & _
        "       c.No As 单据号,c.标识号,c.姓名,c.性别,c.年龄,To_Char(c.登记时间,'YYYY-MM-DD HH24:MI:SS') as 记帐时间" & _
        " From 住院费用记录 c,病人预交记录 d," & _
        "     ( Select d.No" & _
        "       From 病人预交记录 d," & _
        "           (   Select Id As 结帐id From 病人结帐记录 Where 收费时间>=[1] And 收费时间<[2] ) l" & _
        "       Where d.结帐id = l.结帐id And d.记录性质 In (1, 11)" & _
        "       Group By d.No" & _
        "       Having d.No Is Not Null And Sum(d.金额) - Sum(d.冲预交) <> 0 " & _
        "      ) n" & _
        " Where d.No = n.No And d.记录性质 In (1, 11)" & _
        "       And c.结帐ID=d.结帐ID And c.记录性质 IN(2, 12, 3, 13, 5, 15) And c.记帐费用=1"
        
        gstrSQL = gstrSQL & " Union ALL " & _
        " Select Distinct Decode(Mod(c.记录性质,10),2,Decode(c.门诊标志,2,'住院记帐','门诊记帐'),3,'自动记帐',5,'就诊卡记帐') As 单据类型, '结算时使用的预交款未用完' as 无法转移原因," & _
        "       c.No As 单据号,c.标识号,c.姓名,c.性别,c.年龄,To_Char(c.登记时间,'YYYY-MM-DD HH24:MI:SS') as 记帐时间" & _
        " From 门诊费用记录 c,病人预交记录 d," & _
        "     ( Select d.No" & _
        "       From 病人预交记录 d," & _
        "           (   Select Id As 结帐id From 病人结帐记录 Where 收费时间>=[1] And 收费时间<[2] ) l" & _
        "       Where d.结帐id = l.结帐id And d.记录性质 In (1, 11)" & _
        "       Group By d.No" & _
        "       Having d.No Is Not Null And Sum(d.金额) - Sum(d.冲预交) <> 0 " & _
        "      ) n" & _
        " Where d.No = n.No And d.记录性质 In (1, 11)" & _
        "       And c.结帐ID=d.结帐ID And c.记录性质 IN(2, 12, 3, 13, 5, 15) And c.记帐费用=1" & _
        " Order By 单据类型,单据号 Desc"
        
        
    Case 2
        '门诊病人
        gstrSQL = _
        " Select Decode(Count(d.No),0,Null,'病人挂号费用未转出') ||Decode(Count(e.挂号id),0,Null,CHR(13)||CHR(10)||'存在未转出的医嘱费用')  ||Decode(Count(a.挂号id),0,Null,CHR(13)||CHR(10)||'存在未在此之前发送的医嘱') as 无法转移原因," & _
        "       r.No As 挂号单,r.门诊号,r.姓名,r.性别,r.年龄,x.名称 As 就诊科室,To_Char(r.登记时间,'YYYY-MM-DD HH24:MI:SS') As 就诊时间" & _
        " From 部门表 x,病人挂号记录 r," & _
        "   (   Select No From 门诊费用记录 Where 登记时间>=[1] And 登记时间<[2] And 记录性质 = 4) d," & _
        "   (   Select r.Id As 挂号id From 病人医嘱记录 a, 病人挂号记录 r Where a.挂号单 = r.No And r.登记时间>=[1] And r.登记时间<[2] Group By r.Id Having Max(a.停嘱时间)>=[2]) a," & _
        "   (   Select a.挂号id From 门诊费用记录 e, (Select a.Id, r.Id As 挂号id  From 病人医嘱记录 a, 病人挂号记录 r  Where a.挂号单 = r.No And r.登记时间>=[1] And r.登记时间<[2]) a" & _
        "       Where e.医嘱序号 = a.Id) e" & _
        " Where r.No = d.No(+) And r.Id = a.挂号id(+) And r.Id = e.挂号id(+)" & _
        "       And r.执行状态<>2 And r.登记时间>=[1] And r.登记时间<[2] And x.Id=r.执行部门ID" & _
        " Group By r.No,r.门诊号,r.姓名,r.性别,r.年龄,x.名称,r.登记时间" & _
        " Having Count(d.No) > 0 Or Count(a.挂号id) > 0 Or Count(e.挂号id) > 0" & _
        " Order By 就诊时间 Desc,门诊号"
    Case 3
        '住院病人
        gstrSQL = _
            " Select '病人存在未转出费用' as 无法转移原因,i.住院号,i.姓名,i.性别,i.年龄,p.主页id As 住院次数,d.名称 As 住院科室," & _
            "        To_Char(p.入院日期,'YYYY-MM-DD HH24:MI') as 入院时间,To_Char(p.出院日期,'YYYY-MM-DD HH24:MI') as 出院时间" & _
            " From 部门表 d,病人信息 i,病案主页 p" & _
            " Where p.出院日期>=[1] And p.出院日期<[2] And Nvl(p.数据转出, 0) <> 1" & _
                " And i.病人ID=p.病人ID And p.出院科室ID=d.ID" & _
                " And Exists (Select 1 From 住院费用记录 Where 病人id = p.病人id And 主页id = p.主页id)" & _
            " Order BY 出院日期 Desc,住院号"

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Case 4
        '体检任务数据
        gstrSQL = _
            "Select '受检人员的费用未转出' As 无法转移原因, c.姓名, c.性别, c.年龄, c.门诊号, c.健康号, b.任务编号, a.结束时间 As 完成时间, d.名称 As 体检科室" & vbNewLine & _
            "From 体检任务人员 A, 体检任务记录 B, 病人信息 C, 部门表 D" & vbNewLine & _
            "Where a.结束时间 > [1] And a.结束时间 < [2] And a.任务id = b.Id And a.病人id = c.病人id And b.体检部门id = d.Id And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From 门诊费用记录 X, 体检任务费用 Y" & vbNewLine & _
            "       Where y.任务id = a.任务id And y.病人id = a.病人id And x.门诊标志 = 4 And x.No = y.No And x.记录性质 = y.记录性质)" & vbNewLine & _
            "Order By a.结束时间 Desc, c.姓名, c.病人id, b.任务编号"
            
    End Select
    
    Set mrsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(strBegin), CDate(strEnd))
    
    With Me.vsData
        Set .DataSource = mrsData
        Err.Clear
        Call RestoreFlexState(vsData, App.ProductName & "\" & Me.Name & mintType)
        If .Rows = .FixedRows Then
            .Rows = .FixedRows + 1
            cmdExcel.Enabled = False
            cmdFind.Enabled = False
        End If
        If mintType = 2 Then
            .WordWrap = True
            .ColWidth(0) = 1800
            .AutoSizeMode = flexAutoSizeRowHeight
            .AutoSize 0
        End If
        
        cboFind.Clear
        For i = 0 To .Cols - 1
            If InStr("单据号,挂号单,姓名,住院号,门诊号,标识号", .TextMatrix(0, i)) > 0 Then
                cboFind.AddItem .TextMatrix(0, i)
                cboFind.ItemData(cboFind.NewIndex) = i
            End If
            .ColAlignment(i) = 1
        Next
        cboFind.ListIndex = 0
        
        .RowHeight(0) = 250
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
        .Row = 1
    End With
    
    Screen.MousePointer = 0
    LoadData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lblNote_Change()
    mlngBegin = 0
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdFind.Enabled Then Call cmdFind_Click
    ElseIf InStr("单据号,挂号单", cboFind.Text) > 0 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub
