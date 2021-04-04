VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFinaceSuperviseCustomInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "财务缴款登记卡"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8625
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinaceSuperviseCustomInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtTotal 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   350
      Left            =   810
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6000
      Width           =   7665
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "打印设置(&S)"
      Height          =   350
      Left            =   1260
      TabIndex        =   21
      Top             =   7590
      Width           =   1590
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   20
      Top             =   7590
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6015
      TabIndex        =   15
      Top             =   7590
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7215
      TabIndex        =   18
      Top             =   7590
      Width           =   1100
   End
   Begin VB.TextBox txtInputPerson 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Left            =   810
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6885
      Width           =   1785
   End
   Begin VB.TextBox txtTime 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Left            =   5835
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6885
      Width           =   2625
   End
   Begin VB.TextBox txtMemo 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   810
      MaxLength       =   500
      TabIndex        =   10
      Top             =   6450
      Width           =   7665
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6540
      TabIndex        =   5
      Top             =   1230
      Width           =   1935
   End
   Begin VB.ComboBox cboDept 
      Height          =   330
      Left            =   3435
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1222
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.ComboBox cboNO 
      Height          =   330
      Left            =   6540
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   720
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBalance 
      Height          =   4305
      Left            =   120
      TabIndex        =   6
      Top             =   1590
      Width           =   8355
      _cx             =   14737
      _cy             =   7594
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFinaceSuperviseCustomInput.frx":6852
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   975
      TabIndex        =   1
      Top             =   1230
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   194641923
      CurrentDate     =   41520
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "合  计"
      Height          =   210
      Left            =   165
      TabIndex        =   7
      Top             =   6075
      Width           =   630
   End
   Begin VB.Label lblTittle 
      Alignment       =   2  'Center
      Caption         =   "财务缴款登记卡"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   105
      TabIndex        =   19
      Top             =   195
      Width           =   8310
   End
   Begin VB.Line linMain 
      BorderColor     =   &H8000000C&
      X1              =   -30
      X2              =   10410
      Y1              =   7305
      Y2              =   7305
   End
   Begin VB.Label lblInputPerson 
      AutoSize        =   -1  'True
      Caption         =   "登记人"
      Height          =   210
      Left            =   165
      TabIndex        =   11
      Top             =   6945
      Width           =   630
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "登记时间"
      Height          =   210
      Left            =   4920
      TabIndex        =   13
      Top             =   6945
      Width           =   840
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "缴款时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   1297
      Width           =   720
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "摘  要"
      Height          =   210
      Left            =   165
      TabIndex        =   9
      Top             =   6480
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "缴款人"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5910
      TabIndex        =   4
      Top             =   1275
      Width           =   630
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "缴款部门"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2655
      TabIndex        =   2
      Top             =   1282
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "NO"
      Height          =   210
      Left            =   6270
      TabIndex        =   17
      Top             =   765
      Width           =   210
   End
End
Attribute VB_Name = "frmFinaceSuperviseCustomInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mblnOtherPerson As Boolean
Private mstr缴款人 As String, mlng缴款人ID  As Long
Private mrsBalance As ADODB.Recordset
Private mblnChange As Boolean '是否被用户操作过
Private mblnSuccess As Boolean
Private mblnFirst  As Boolean
Public Function EditCard(ByVal frmMain As Object, _
    ByVal str缴款人 As String, ByVal lng缴款人ID As Long, _
    ByVal lngModule As Long, ByVal strPrivs As String, Optional ByVal blnOtherPerson As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序编辑入口(手工缴款)
    '入参:str缴款人-缴款人
    '       lng缴款人ID-缴款人ID
    '       blnOtherPerson-true时为其他人员收款;否则为收费员款款
    '出参:
    '返回:保存成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-11 18:08:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mstr缴款人 = str缴款人: mlng缴款人ID = lng缴款人ID
    mlngModule = lngModule: mstrPrivs = strPrivs
    mblnOtherPerson = blnOtherPerson
    Call InitFace
    If LoadCollectData = False Then mblnChange = False: Unload Me: Exit Function
    mblnChange = False: mblnSuccess = False
    If frmMain Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmMain
    End If
    EditCard = mblnSuccess
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面信息
    '编制:刘兴洪
    '日期:2013-10-11 16:00:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim datCurrnet  As Date
    txtName.Text = mstr缴款人: txtInputPerson.Text = UserInfo.姓名
    
    datCurrnet = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    dtpDate.Value = datCurrnet
    dtpDate.MaxDate = datCurrnet
    
    Call InitGrid
    'Call LoadDept
End Sub
Private Function LoadCollectData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收款数据
    '返回:加载成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-11 18:14:31
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long
    On Error GoTo errHandle
        
    strSQL = "" & _
    "   Select decode(nvl(M.性质,0),1,1,2,2,3,10,4,11,4) as 序号, A.结算方式,A.余额 " & _
    "   From 人员缴款余额 A,结算方式 M" & _
    "   Where A.结算方式=M.名称(+)  and A.性质=1 and nvl(A.余额,0)<>0  " & _
    "           And  A.收款员 =[1] " & _
    "   Order by 序号,结算方式"
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人)
    If mrsBalance.RecordCount = 0 Then
        MsgBox "收费员『" & mstr缴款人 & "』没有暂存金额，无须进行缴款操作。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    With vsBalance
        .Clear 1: .Rows = IIf(mrsBalance.RecordCount = 0, 1, mrsBalance.RecordCount) + 1
        i = 1
        Do While Not mrsBalance.EOF
             .TextMatrix(i, .ColIndex("序号")) = NVL(mrsBalance!序号)
             .TextMatrix(i, .ColIndex("结算方式")) = NVL(mrsBalance!结算方式)
             .Cell(flexcpData, i, .ColIndex("结算方式")) = Trim(NVL(mrsBalance!结算方式))
             .TextMatrix(i, .ColIndex("金额")) = Format(Val(NVL(mrsBalance!余额)), "###0.00;-###0.00;0.00;0.00")
             .TextMatrix(i, .ColIndex("结算号码")) = ""
            i = i + 1
            mrsBalance.MoveNext
        Loop
        .ColComboList(.ColIndex("结算方式")) = .BuildComboList(mrsBalance, "结算方式,余额", "结算方式")
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Name, "结算方式列表", False
    End With
    Call CalcTotal
    LoadCollectData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格
    '编制:刘兴洪
    '日期:2013-10-11 15:59:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsBalance
           .Clear 1
           .Cols = 4: .Rows = 2
           .FixedRows = 1
           .TextMatrix(0, 0) = "序号"
           .TextMatrix(0, 1) = "结算方式"
           .TextMatrix(0, 2) = "金额"
           .TextMatrix(0, 3) = "结算号码"
           For i = 0 To .Cols - 1
               .ColKey(i) = .TextMatrix(0, i)
               If i = .ColIndex("金额") Then
                   .ColAlignment(i) = flexAlignRightCenter
               Else
                   .ColAlignment(i) = flexAlignLeftCenter
               End If
               .FixedAlignment(i) = flexAlignCenterCenter
           Next
           .ColHidden(.ColIndex("序号")) = True
           .ExtendLastCol = True
           .AutoSizeMode = flexAutoSizeColWidth
           .AutoResize = True
           Call .AutoSize(0, .Cols - 1)
           zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Name, "结算方式列表", False
           .Editable = flexEDKbdMouse
    End With
End Sub
Private Sub CalcTotal()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算缴款总金额
    '编制:刘兴洪
    '日期:2013-10-11 16:14:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTemp As Double, i As Integer
    With vsBalance
        For i = 1 To .Rows - 1
             dblTemp = dblTemp + Val(.TextMatrix(i, .ColIndex("金额")))
        Next
    End With
    txtTotal.Text = Format(dblTemp, "###0.00;-###0.00;0;") & "元" & IIf(dblTemp = 0, "", " （" & zlCommFun.UppeMoney(dblTemp) & "）")
End Sub

Private Function LoadDept() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载缴款人部门信息
    '编制:刘兴洪
    '日期:2013-09-11 14:05:08
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
        
    strSQL = "" & _
    "   Select Distinct a.Id, a.编码, a.名称,b.缺省" & vbNewLine & _
    "   From 部门表 a, 部门人员 b" & vbNewLine & _
    "   Where a.Id = b.部门id And b.人员ID=[1] " & vbNewLine & _
     "              And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
    "               And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
    "   Order By a.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng缴款人ID)
    With cboDept
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!编码) & "-" & rsTemp!名称
            .ItemData(.NewIndex) = Val(NVL(rsTemp!ID))
            If Val(NVL(rsTemp!缺省)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 And .ListCount <> 0 Then .ListIndex = 0
    End With
    LoadDept = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'Private Sub cboDept_Click()
'    mblnChange = True
'End Sub

'Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
'
'End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call cmdHelp_Click
End Sub

Private Sub Form_Load()
    mblnFirst = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strNO As String
    On Error GoTo errHandle
    If isValied() = False Then Exit Sub
    If SaveData(strNO) = False Then Exit Sub
    mblnChange = False: mblnSuccess = True
    Call BillPrint(strNO)
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdPrintSet_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me
End Sub
Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txtMemo_Change()
    mblnChange = True
End Sub
Private Sub txtMemo_GotFocus()
    zlControl.TxtSelAll txtMemo
    zlCommFun.OpenIme True
End Sub
Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txtMemo_LostFocus()
    zlCommFun.OpenIme False
End Sub
Private Sub vsBalance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblMoney As Double
    With vsBalance
        mblnChange = True
        Select Case Col
        Case .ColIndex("金额")
            Call CalcTotal '重新计算总金额
        Case .ColIndex("结算号码")
        Case .ColIndex("结算方式")
            dblMoney = 0
            If Not mrsBalance Is Nothing Then
                mrsBalance.Filter = "结算方式='" & .TextMatrix(Row, Col) & "'"
                If Not mrsBalance.EOF Then
                    dblMoney = Val(NVL(mrsBalance!余额))
                End If
            End If
            .TextMatrix(Row, .ColIndex("金额")) = Format(dblMoney, "##0.00;-##0.00;0.00;")
            Call CalcTotal '重新计算总金额
        End Select
    End With
End Sub
Private Sub vsBalance_GotFocus()
    Call zl_VsGridGotFocus(vsBalance)
    zlCommFun.OpenIme False
End Sub
Private Sub vsBalance_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsBalance)
End Sub
Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "结算方式列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsBalance, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "结算方式列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBalance
        Select Case Col
        Case .ColIndex("金额"), .ColIndex("结算号码"), .ColIndex("结算方式")
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub
Private Sub vsBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsBalance
        If .Col = .Cols - 1 And .Row = .Rows - 1 _
            And Trim(.TextMatrix(.Row, .ColIndex("结算方式"))) = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    End With
    Call zlVsMoveGridCell(vsBalance, vsBalance.ColIndex("结算方式"), vsBalance.Cols - 1, True)
End Sub
Private Sub vsBalance_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call zlVsMoveGridCell(vsBalance, vsBalance.ColIndex("结算方式"), vsBalance.Cols - 1, True)
End Sub
Private Sub vsBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub
Private Sub vsBalance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsBalance
        If Row <= 1 Then Exit Sub
        Select Case Col
        Case .ColIndex("结算号码")
            VsFlxGridCheckKeyPress vsBalance, Row, Col, KeyAscii, m文本式
            If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Or KeyAscii = Asc(",") Then KeyAscii = 0: Exit Sub
        Case .ColIndex("金额")
            VsFlxGridCheckKeyPress vsBalance, Row, Col, KeyAscii, m负金额式
        End Select
    End With
End Sub
Private Sub vsBalance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer
    '数据验证
    With vsBalance
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
        Case .ColIndex("结算号码")
            If zlCommFun.ActualLen(strKey) > 10 Then
                MsgBox "结算号码超长,最多只能输入10个字符或5个汉字", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            If InStr(1, strKey, "'") > 0 Or InStr(1, strKey, "|") > 0 Or InStr(1, strKey, ",") > 0 Then
                MsgBox "结算号码中不能包含特殊字符:',| ", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
        Case .ColIndex("金额")
            If Not IsNumeric(strKey) Then
                MsgBox "金额必须输入数字,不能输入其他字符。", vbInformation, gstrSysName
                Cancel = True: Exit Sub
             End If
             If Val(strKey) > 999999999 Then
                MsgBox "金额输入过大,最大只能输入999999999。", vbInformation, gstrSysName
                Cancel = True: Exit Sub
             End If
             If Val(strKey) < -999999999 Then
                MsgBox "金额输入过小,最大只能输入-999999999。", vbInformation, gstrSysName
                Cancel = True: Exit Sub
                Exit Sub
             End If
             .EditText = Format(strKey, "###0.00;-###0.00;0.00;0.00")
        End Select
    End With
End Sub



Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据合法性检查
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-11 16:35:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblMoney As Double, strTemp As String, j As Long
    Dim str结算方式 As String
    On Error GoTo errHandle
    isValied = False
    '问题号:110281,焦博,2017/08/15,把轧账说明的上限从50个字符调整为500个字符
    If zlCommFun.ActualLen(txtMemo.Text) > 500 Then
        MsgBox "摘要的长度不能超过250个汉字或500个字符。", vbInformation, gstrSysName
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
    If InStr(txtMemo.Text, "'") > 0 Then
        MsgBox "摘要含有非法字符（'）。", vbInformation, gstrSysName
        zlControl.TxtSelAll txtMemo
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
'    If cboDept.ListIndex < 0 Then
'        MsgBox "未选择缴款部门!", vbInformation, gstrSysName
'        If cboDept.Visible And cboDept.Enabled Then cboDept.SetFocus
'        Exit Function
'    End If
    With vsBalance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
           If str结算方式 <> "" Then
                strTemp = .TextMatrix(i, .ColIndex("结算号码"))
                If zlCommFun.ActualLen(strTemp) > 10 Then
                    MsgBox "结算号码超长,最多只能输入10个字符或5个汉字", vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("结算号码")
                    If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                        .TopRow = .Row: .LeftCol = .Col
                    End If
                    If .Visible And .Enabled Then .SetFocus
                    Exit Function
                End If
                If InStr(1, strTemp, "'") > 0 Or InStr(1, strTemp, "|") > 0 Or InStr(1, strTemp, ",") > 0 Then
                    MsgBox "结算号码中不能包含特殊字符:',| ", vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("结算号码")
                    If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                        .TopRow = .Row: .LeftCol = .Col
                    End If
                    If .Visible And .Enabled Then .SetFocus
                    Exit Function
                End If
                strTemp = Trim(.TextMatrix(i, .ColIndex("金额")))
                If Not IsNumeric(strTemp) Then
                   MsgBox "金额必须输入数字,不能输入其他字符。", vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("金额")
                    If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                        .TopRow = .Row: .LeftCol = .Col
                    End If
                    If .Visible And .Enabled Then .SetFocus
                    Exit Function
                End If
                If Val(strTemp) > 999999999 Then
                   MsgBox "金额输入过大,最大只能输入999999999。", vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("金额")
                    If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                        .TopRow = .Row: .LeftCol = .Col
                    End If
                    If .Visible And .Enabled Then .SetFocus
                    Exit Function
                End If
                If Val(strTemp) < -999999999 Then
                   MsgBox "金额输入过小,最大只能输入-999999999。", vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("金额")
                    If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                        .TopRow = .Row: .LeftCol = .Col
                    End If
                    If .Visible And .Enabled Then .SetFocus
                    Exit Function
                End If
                mrsBalance.Filter = "结算方式='" & str结算方式 & "'"
                If mrsBalance.EOF Then
                    If MsgBox("缴款人不存在" & str结算方式 & "的暂存金,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        .Row = i: .Col = .ColIndex("金额")
                        If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                            .TopRow = .Row: .LeftCol = .Col
                        End If
                        If .Visible And .Enabled Then .SetFocus
                        Exit Function
                    End If
                Else
                    '检查金额是否正确
                    If Val(strTemp) > Val(NVL(mrsBalance!余额)) Then
                        If MsgBox(str结算方式 & "的缴款金额(" & Format(Val(strTemp), "0.00") & ")大于暂存金额(" & Format(Val(NVL(mrsBalance!余额)), "0.00") & ")，是否继续？", vbYesNo Or vbQuestion Or vbDefaultButton2, Me.Caption) = vbNo Then
                            .Row = i: .Col = .ColIndex("金额")
                            If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                                .TopRow = .Row: .LeftCol = .Col
                            End If
                            If .Visible And .Enabled Then .SetFocus
                            Exit Function
                        End If
                    End If
                End If
                
                '检查结算方式是否重复
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("结算方式"))) = Trim(.TextMatrix(j, .ColIndex("结算方式"))) And i <> j Then
                        MsgBox "第" & i & "行与第" & j & "行的结算方式相同,请合并。", vbInformation, gstrSysName
                         .Row = i: .Col = .ColIndex("金额")
                         If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                             .TopRow = .Row: .LeftCol = .Col
                         End If
                         If .Visible And .Enabled Then .SetFocus
                         Exit Function
                    End If
                Next
            End If
        Next
    End With
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData(ByRef strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '出参:strNo-数据保存成功后,返回成功的单据号
    '返回:保存成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-11 16:59:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lngID As Long, strTemp As String, i As Long
    Dim str结算方式 As String, str结算金额 As String, str结算号码 As String
 
    On Error GoTo errHandle
    With vsBalance
        For i = 1 To .Rows - 1
            strTemp = .TextMatrix(i, .ColIndex("结算方式"))
            If strTemp <> "" Then
                str结算方式 = str结算方式 & "," & strTemp
                str结算金额 = str结算金额 & "," & Val(.TextMatrix(i, .ColIndex("金额")))
                str结算号码 = str结算号码 & "," & Trim(.TextMatrix(i, .ColIndex("结算号码")))
            End If
        Next
    End With
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
    If str结算金额 <> "" Then str结算金额 = Mid(str结算金额, 2)
    If str结算号码 <> "" Then str结算号码 = Mid(str结算号码, 2)
    
    If str结算方式 = "" Then
        MsgBox "不存在缴款数据,你必须输入缴款数据,才能进行正常收款", vbInformation + vbOKOnly, gstrSysName
        If vsBalance.Enabled And vsBalance.Visible Then vsBalance.SetFocus
        Exit Function
    End If
    
    If zlCommFun.ActualLen(str结算方式) > 4000 Then
        MsgBox "在结算明细信息中输入的结算方式过多,不能进行收款", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If zlCommFun.ActualLen(str结算金额) > 4000 Then
        MsgBox "在结算明细信息中输入的结算方式所对应的结算金额过多,不能进行收款", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If zlCommFun.ActualLen(str结算号码) > 4000 Then
        MsgBox "在结算明细信息中输入的结算方式所对应的结算号码过多,不能进行收款", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If str结算号码 = "" Then str结算号码 = " "
    lngID = zlDatabase.GetNextId("人员收缴记录")
    strNO = zlDatabase.GetNextNo(140)
    'Zl_手工收款记录_Insert
    strSQL = "Zl_手工收款记录_Insert("
    '  Id_In         In 人员收缴记录.Id%Type,
    strSQL = strSQL & "" & lngID & ","
    '  No_In         In 人员收缴记录.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  收款员_In     In 人员收缴记录.收款员%Type,
    strSQL = strSQL & "'" & mstr缴款人 & "',"
    '  收款部门id_In In 人员收缴记录.收款部门id%Type,
    strSQL = strSQL & "" & "Null,"
    'strSQL = strSQL & "" & cboDept.ItemData(cboDept.ListIndex) & ","
    '  收款时间_In   In 人员收缴记录.开始时间%Type,
    strSQL = strSQL & "to_date('" & Format(dtpDate.Value, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
    '  摘要_In       In 人员收缴记录.摘要%Type,
    strSQL = strSQL & IIf(Trim(txtMemo.Text) = "", "NULL", "'" & txtMemo.Text & "'") & ","
    '  登记人_In     In 人员收缴记录.登记人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  登记时间_In   In 人员收缴记录.登记时间%Type,
    strSQL = strSQL & "Sysdate,"
    ' 收缴标志_In   In 人员收缴记录.收缴标志%Type,
    strSQL = strSQL & IIf(mblnOtherPerson, 1, "NULL") & ","
    '  结算方式_In   Varchar2,结算方式_IN:允许多个,多个时,用逗号分离,比如:现金,支票,...
    '       结算方式_In,结算金额_In,结算号码_IN 三个参数的值的个数要一一对应:比如:结算方式_IN (现金,支票...),结算金额(100,0...),结算号码_IN(A001,A002,...)
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  结算金额_In   Varchar2,结算金额_IN:允许多个,多个时,用逗号分离,与结算方式_IN 一一对应.
    strSQL = strSQL & "'" & str结算金额 & "',"
    '  结算号码_In   In Varchar2,结算号码_In:允许多个,多个时,用逗号分离,与结算方式_IN 一一对应
    strSQL = strSQL & "'" & str结算号码 & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub BillPrint(ByVal strNO As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收款收据打印
    '编制:刘兴洪
    '日期:2013-09-11 11:55:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    blnPrint = False
    If Not zlStr.IsHavePrivs(mstrPrivs, "收款收据打印") Then Exit Sub
    Select Case Val(zlDatabase.GetPara("收款收据打印方式", glngSys, mlngModule))     '使用医生站的相关参数
    Case 0    '不打印
        Exit Sub
    Case 1    '自助动打印
        blnPrint = True
    Case 2    '选择打印
        If MsgBox("你是否要打印缴款收据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            blnPrint = True
        End If
    End Select
    If blnPrint = False Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500", Me, "NO=" & strNO, "记录性质=5", 2)
End Sub
