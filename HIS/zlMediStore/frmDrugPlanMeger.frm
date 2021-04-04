VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmDrugPlanMeger 
   Caption         =   "计划单合并"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10350
   Icon            =   "frmDrugPlanMeger.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   10350
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic基本条件 
      Height          =   3120
      Left            =   60
      ScaleHeight     =   3060
      ScaleWidth      =   3435
      TabIndex        =   12
      Top             =   1125
      Width           =   3500
      Begin VB.CommandButton cmdFind 
         Caption         =   "过滤(&O)"
         Height          =   350
         Left            =   2040
         TabIndex        =   6
         Top             =   2370
         Width           =   1100
      End
      Begin VB.TextBox txt结束No 
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   540
         Width           =   2200
      End
      Begin VB.TextBox txt开始No 
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   135
         Width           =   2200
      End
      Begin VB.ComboBox cbo时间范围 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   2200
      End
      Begin MSComCtl2.DTPicker DTP结束时间 
         Height          =   300
         Left            =   960
         TabIndex        =   5
         Top             =   1800
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   120782851
         CurrentDate     =   40750
      End
      Begin MSComCtl2.DTPicker DTP开始时间 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   1380
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   120782851
         CurrentDate     =   40750
      End
      Begin VB.Label lbl结束No 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "结束NO"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lbl开始No 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "开始NO"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lbl时间范围 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "复核时间"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lbl开始时间 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lbl结束时间 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   1860
         Width           =   720
      End
   End
   Begin VB.Frame fra基本条件 
      Height          =   495
      Left            =   60
      TabIndex        =   10
      Top             =   495
      Width           =   3500
      Begin VB.Label lbl基本条件 
         AutoSize        =   -1  'True
         Caption         =   "基本条件"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   195
         Width           =   720
      End
   End
   Begin VB.Frame fra计划单信息 
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   495
      Width           =   6270
      Begin VB.CheckBox chkAllSelect 
         Caption         =   "全选"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1425
         TabIndex        =   18
         Top             =   165
         Width           =   975
      End
      Begin VB.Label lbl计划单信息 
         AutoSize        =   -1  'True
         Caption         =   "计划单信息"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   195
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1515
      Left            =   3840
      TabIndex        =   7
      Top             =   1230
      Width           =   6255
      _cx             =   11033
      _cy             =   2672
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
      BackColorSel    =   16777152
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugPlanMeger.frx":030A
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.Frame fraEW 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4185
      Left            =   3660
      MousePointer    =   9  'Size W E
      TabIndex        =   9
      Top             =   465
      Width           =   45
   End
   Begin XtremeCommandBars.CommandBars comBars 
      Left            =   315
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgIcon 
      Bindings        =   "frmDrugPlanMeger.frx":037F
      Left            =   1515
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDrugPlanMeger.frx":0393
   End
End
Attribute VB_Name = "frmDrugPlanMeger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnChkFocus As Boolean
Private mblnChang    As Boolean

Private Const MCONMEGER = 2 '合并
Private Const MCONEXIT = 4  '退出

Private Const MCON选择 = 0
Private Const MCONNO = 1
Private Const MCONID = 2
Private Const MCON计划类型 = 3
Private Const MCON期间 = 4
Private Const MCON所属库房 = 5
Private Const MCON编制方法 = 6
Private Const MCON编制人 = 7
Private Const MCON编制日期 = 8
Private Const MCON审核人 = 9
Private Const MCON审核日期 = 10
Private Const MCON复核人 = 11
Private Const MCON复核日期 = 12
Private Const MCON编制说明 = 13
Private Const MCONROWS = 14

Private Sub GetList()
'提取计划单数据
    Dim cmdControl As CommandBarControl
    Dim rsList As New Recordset
    Dim strFind As String
    Dim lngRow As Long
     
    On Error GoTo errHandle
    
    With vsfList
        If Me.txt开始No <> "" And Me.txt结束No <> "" Then strFind = " And A.No >= [3] And A.No <=[4] "
        If Me.txt开始No <> "" And Me.txt结束No = "" Then strFind = " And A.No >= [3] "
        If Me.txt开始No = "" And Me.txt结束No <> "" Then strFind = " And A.No <= [4] "
        
        gstrSQL = " SELECT '' as 选择, a.NO, a.ID, DECODE(a.计划类型,0,'临时',1,'月度计划',2,'季度计划',3,'年度计划','周计划') AS 计划类型 ," & _
            "a.期间, b.名称 as 所属库房, DECODE(A.编制方法, 0, '根据申领产生', 1, '往年同期线形参照法', 2, '临近期间平均参照法', 3, '药品储备定额参照法', 4, '药品日销售量参照法', '自定义区间参照法') AS 编制方法 ," & _
            "a.编制人,TO_CHAR(a.编制日期,'YYYY-MM-DD HH24:MI:SS') AS 编制日期, a.审核人, " & _
            "TO_CHAR(a.审核日期,'YYYY-MM-DD HH24:MI:SS') AS 审核日期,a.复核人,TO_CHAR(a.复核日期,'YYYY-MM-DD HH24:MI:SS') AS 复核日期, a.编制说明 " & _
            " FROM 药品采购计划 A, 部门表 B " & _
            " WHERE a.库房ID= b.ID(+) And NVL(a.药房id,0)=0 And NVL(a.库房ID,0)<>0 And NVL(a.合并计划id,0)=0 And a.复核日期 between [1] And to_date([2],'YYYY-MM-DD HH24:MI:SS') " & strFind & _
            " ORDER BY A.NO DESC "
            
        Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, "合并计划单", DTP开始时间.Value, Format(DTP结束时间.Value, "yyyy-mm-dd") & " 23:59:59", txt开始No.Text, txt结束No.Text)
        .rows = 1
        
        If rsList.EOF Then
            .rows = 2
            .Editable = flexEDNone
            chkAllSelect.Enabled = False
        Else
            .rows = rsList.RecordCount + 1
            .Editable = flexEDKbdMouse
            chkAllSelect.Enabled = True
            For lngRow = 1 To .rows - 1
                .TextMatrix(lngRow, .ColIndex("NO")) = rsList!NO
                .TextMatrix(lngRow, .ColIndex("ID")) = rsList!id
                .TextMatrix(lngRow, .ColIndex("计划类型")) = rsList!计划类型
                .TextMatrix(lngRow, .ColIndex("期间")) = rsList!期间
                .TextMatrix(lngRow, .ColIndex("所属库房")) = rsList!所属库房
                .TextMatrix(lngRow, .ColIndex("编制方法")) = rsList!编制方法
                .TextMatrix(lngRow, .ColIndex("编制人")) = rsList!编制人
                .TextMatrix(lngRow, .ColIndex("编制日期")) = rsList!编制日期
                .TextMatrix(lngRow, .ColIndex("审核人")) = rsList!审核人
                .TextMatrix(lngRow, .ColIndex("审核日期")) = rsList!审核日期
                .TextMatrix(lngRow, .ColIndex("复核人")) = rsList!复核人
                .TextMatrix(lngRow, .ColIndex("复核日期")) = rsList!复核日期
                .TextMatrix(lngRow, .ColIndex("编制说明")) = IIf(IsNull(rsList!编制说明), "", rsList!编制说明)
                rsList.MoveNext
            Next
        End If
        
        .Row = 1
        .SetFocus
        chkAllSelect.Value = 0
        rsList.Close
    End With
    
    Set cmdControl = comBars.FindControl(, MCONMEGER)
    cmdControl.Enabled = False
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub MegerPlan()
'合并计划单
    Dim intRow As Integer
    Dim strID As String
    Dim strNewNO As String
    Dim lngNewId As Long
    Dim arrSql As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    arrSql = Array()
    strNewNO = Sys.GetNextNo(32)
    lngNewId = Sys.NextId("药品采购计划")
    strID = lngNewId & "|"
    With vsfList
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("选择"))) = -1 Then
                strID = strID & .TextMatrix(intRow, .ColIndex("id")) & ","
            End If
        Next
    End With
    
    gstrSQL = "zl_药品计划管理主表_INSERT("
        '计划ID
        gstrSQL = gstrSQL & lngNewId
        'NO
        gstrSQL = gstrSQL & ",'" & strNewNO & "'"
        '计划类型
        gstrSQL = gstrSQL & ",1"
        '期间
        gstrSQL = gstrSQL & ",'" & Format(DateAdd("m", 1, Sys.Currentdate), "yyyyMM") & "'"
        '库房ID
        gstrSQL = gstrSQL & ",Null"
        '药房ID
        gstrSQL = gstrSQL & ",Null"
        '编制方法
        gstrSQL = gstrSQL & ",1"
        '编制人
        gstrSQL = gstrSQL & ",'" & UserInfo.用户姓名 & "'"
        '编制日期
        gstrSQL = gstrSQL & ",to_date('" & Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
        '编制说明
        gstrSQL = gstrSQL & ",''"
        '来源库房
        gstrSQL = gstrSQL & ",'0'"
        gstrSQL = gstrSQL & ")"
        
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSQL
    
    gstrSQL = "Zl_药品计划管理_Union('" & Mid(strID, 1, Len(strID) - 1) & "')"
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSQL
    
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "MegerPlan")
    Next
    gcnOracle.CommitTrans
    
    '重新获取数据
    Call GetList
    MsgBox "合并成功！", vbInformation, gstrSysName
    
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetListCol()
'表格属性设置
    Dim intCol As Integer
    
    With vsfList
        .rows = 2
        .Cols = MCONROWS
        .ColDataType(0) = flexDTBoolean
        .Editable = flexEDNone
        .TextMatrix(0, MCON选择) = "选择"
        .TextMatrix(0, MCONNO) = "NO"
        .TextMatrix(0, MCONID) = "ID"
        .TextMatrix(0, MCON计划类型) = "计划类型"
        .TextMatrix(0, MCON期间) = "期间"
        .TextMatrix(0, MCON所属库房) = "所属库房"
        .TextMatrix(0, MCON编制方法) = "编制方法"
        .TextMatrix(0, MCON编制人) = "编制人"
        .TextMatrix(0, MCON编制日期) = "编制日期"
        .TextMatrix(0, MCON审核人) = "审核人"
        .TextMatrix(0, MCON审核日期) = "审核日期"
        .TextMatrix(0, MCON复核人) = "复核人"
        .TextMatrix(0, MCON复核日期) = "复核日期"
        .TextMatrix(0, MCON编制说明) = "编制说明"
        
        For intCol = 0 To .Cols - 1
            .ColKey(intCol) = .TextMatrix(0, intCol)
        Next
        .ColWidth(.ColIndex("选择")) = 500
        .ColWidth(.ColIndex("所属库房")) = 1200
        .ColWidth(.ColIndex("编制方法")) = 1800
        .ColWidth(.ColIndex("编制日期")) = 1000
        .ColWidth(.ColIndex("审核日期")) = 1000
        .ColWidth(.ColIndex("复核日期")) = 1000
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("NO")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("期间")) = flexAlignLeftCenter
        .ColWidth(.ColIndex("ID")) = 0
        
    End With
End Sub

Public Sub showMe(ByVal frmPar As Form)
    Me.Show vbModal, frmPar
End Sub

Private Sub chkAllSelect_Click()
    Dim i As Integer
    
    With vsfList
        If mblnChkFocus Then
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("选择")) = IIf(chkAllSelect.Value = 1, -1, 0)
            Next
        End If
    End With
End Sub

Private Sub chkAllSelect_GotFocus()
    mblnChkFocus = True
End Sub

Private Sub chkAllSelect_LostFocus()
    mblnChkFocus = False
End Sub

Private Sub cmdFind_Click()
    Call GetList
End Sub

Private Sub combars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case MCONMEGER
            Call MegerPlan
        Case MCONEXIT
            Call CheckUnLoad
    End Select
End Sub

Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, 32)
        End If
        Me.txt结束No.SetFocus
    End If
End Sub

Private Sub txt开始No_LostFocus()
    If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
        txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, 32)
    End If
End Sub

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt结束No) < 8 And Len(txt结束No) > 0 Then
            txt结束No.Text = zlCommFun.GetFullNO(txt结束No.Text, 32)
        End If
        Me.cbo时间范围.SetFocus
    End If
End Sub

Private Sub txt结束No_LostFocus()
    If Len(txt结束No) < 8 And Len(txt结束No) > 0 Then
        txt结束No.Text = zlCommFun.GetFullNO(txt结束No.Text, 32)
    End If
End Sub

Private Sub cbo时间范围_Click()
    If Me.cbo时间范围.ListIndex = 0 Then
        Me.DTP开始时间.Value = Date
        Me.DTP结束时间.Value = Date
        Me.DTP开始时间.Enabled = False
        Me.DTP结束时间.Enabled = False
    ElseIf Me.cbo时间范围.ListIndex = 1 Then
        Me.DTP开始时间.Value = Date - 1
        Me.DTP结束时间.Value = Date
        Me.DTP开始时间.Enabled = False
        Me.DTP结束时间.Enabled = False
    ElseIf Me.cbo时间范围.ListIndex = 2 Then
        Me.DTP开始时间.Value = Date - 2
        Me.DTP结束时间.Value = Date
        Me.DTP开始时间.Enabled = False
        Me.DTP结束时间.Enabled = False
    ElseIf Me.cbo时间范围.ListIndex = 3 Then
        Me.DTP开始时间.Value = Date - 6
        Me.DTP结束时间.Value = Date
        Me.DTP开始时间.Enabled = False
        Me.DTP结束时间.Enabled = False
    Else
        Me.DTP开始时间.Value = Date - 30
        Me.DTP结束时间.Value = Date
        Me.DTP开始时间.Enabled = True
        Me.DTP结束时间.Enabled = True
    End If
End Sub

Private Sub cbo时间范围_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub dtp开始时间_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub dtp结束时间_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Call InitComman
    Call InitTool
    Call InitCbo
    Call SetListCol
End Sub

Private Sub InitCbo()
'设置下拉框
    With cbo时间范围
        .Clear
        .AddItem "一天内"
        .AddItem "两天内"
        .AddItem "三天内"
        .AddItem "一周内"
        .AddItem "自定义时间"
        .ListIndex = 0
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fra基本条件.Move 10, , 3500
    fraEW.Move fra基本条件.Left + fra基本条件.Width, fra基本条件.Top, 45, Me.ScaleHeight
    fra计划单信息.Move fraEW.Left + fraEW.Width, fra基本条件.Top, Me.ScaleWidth - fraEW.Left - fraEW.Width - 20
    pic基本条件.Move fra基本条件.Left, fra基本条件.Top + fra基本条件.Height + 10, fra基本条件.Width, Me.ScaleHeight - fra基本条件.Top - fra基本条件.Height - 55
    vsfList.Move fra计划单信息.Left, fra计划单信息.Top + fra计划单信息.Height + 10, fra计划单信息.Width, Me.ScaleHeight - fra计划单信息.Top - fra计划单信息.Height - 50
End Sub

Private Sub fraEW_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'------------------------------------------
'条件区和单据区的拉动
'------------------------------------------
    On Error Resume Next
    If Me.fra基本条件.Width + x < 3000 Or Me.vsfList.Width - x < 4000 Then
        Exit Sub
    End If
    
    If Button = 1 Then
        Me.fraEW.Move Me.fraEW.Left + x, Me.fraEW.Top, Me.fraEW.Width, Me.fraEW.Height
        Me.fra基本条件.Move Me.fra基本条件.Left, Me.fra基本条件.Top, Me.fra基本条件.Width + x, Me.fra基本条件.Height
        Me.fra计划单信息.Move Me.fra计划单信息.Left + x, Me.fra计划单信息.Top, Me.fra计划单信息.Width - x, Me.fra计划单信息.Height
        
        Me.pic基本条件.Move Me.pic基本条件.Left, Me.pic基本条件.Top, Me.pic基本条件.Width + x, Me.pic基本条件.Height
        Me.vsfList.Move Me.vsfList.Left + x, Me.vsfList.Top, Me.vsfList.Width - x, Me.vsfList.Height
        Me.cmdFind.Move cmdFind.Left + x
        
        Me.txt结束No.Width = Me.txt结束No.Width + x
        Me.txt开始No.Width = Me.txt开始No.Width + x
        Me.cbo时间范围.Width = Me.cbo时间范围.Width + x
        Me.DTP结束时间.Width = Me.DTP结束时间.Width + x
        Me.DTP开始时间.Width = Me.DTP开始时间.Width + x
    End If
End Sub

Private Sub InitComman()
'--------------------------------------
'初始化CommandBars控件
'--------------------------------------
    With CommandBarsGlobalSettings
        Set .App = App
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '设置中文语言资源文件
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '控件整体的颜色方案，根据系统自动识别
    End With

    With comBars.Options
        .ShowExpandButtonAlways = False '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True '显示按钮提示
        .AlwaysShowFullMenus = False '不常用的菜单项先隐藏
        .UseFadedIcons = True '图标显示为褪色效果
        .IconsWithShadow = True '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True '工具栏显示为大图标
        .SetIconSize True, 24, 24 '设置大图标的尺寸
        .SetIconSize False, 16, 16 '设置小图标的尺寸
    End With

    With comBars
        .VisualTheme = xtpThemeOffice2003 '设置控件显示风格
        .EnableCustomization False '是否允许自定义设置
        .Item(1).Delete
        .Icons = imgIcon.Icons
    End With
End Sub

Private Sub InitTool()
'-----------------------------------------------------
'设置工具栏
'----------------------------------------------------
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    Set objBar = comBars.Add("工具栏1", xtpBarTop)
    objBar.ContextMenuPresent = False '工具栏上点击鼠标右键时不弹出设置菜单
    objBar.ShowTextBelowIcons = False '工具栏中的按钮文字显示在图标右侧
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, MCONMEGER, "合并")
        objControl.Style = xtpButtonIconAndCaption
        objControl.Enabled = False
        Set objControl = .Add(xtpControlButton, MCONEXIT, "退出")
        objControl.Style = xtpButtonIconAndCaption
        objControl.BeginGroup = True
    End With
End Sub

Private Sub vsfList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList
        If Col <> 0 Then
            Cancel = True
        End If
        mblnChkFocus = False
    End With
End Sub

Private Sub vsfList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim cmdControl As CommandBarControl
    Dim intRow As Integer
    Dim intCount As Integer
    
    With vsfList
        If Row > 0 And Val(.TextMatrix(Row, 2)) <> 0 And Col = 0 Then
            For intRow = 1 To .rows - 1
                If Val(.TextMatrix(intRow, 0)) = -1 Then
                    intCount = intCount + 1
                End If
            Next
            
            '是否选择两项以上
            Set cmdControl = comBars.FindControl(, MCONMEGER)
            If intCount >= 2 Then
                cmdControl.Enabled = True
            Else
                cmdControl.Enabled = False
            End If
            
            '是否全选
            If mblnChkFocus = False Then
                If intCount = .rows - 1 Then
                    chkAllSelect.Value = 1
                Else
                    chkAllSelect.Value = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfList_ChangeEdit()
    mblnChang = True
End Sub

Private Sub CheckUnLoad()
'退出前检查是否有选中的单据
    Dim intRow As Integer
    Dim blnChanged As Boolean
    
    blnChanged = True
    With vsfList
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, 0)) = -1 Then
                If MsgBox("存在已选中的计划单还未合并，是否确定退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                    blnChanged = False
                End If
                Exit For
            End If
        Next
            
        If blnChanged Then Unload Me
    End With
End Sub

