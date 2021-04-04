VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatiPressMoney 
   Caption         =   "病人催款管理"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   Icon            =   "frmPatiPressMoney.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11775
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   11775
      TabIndex        =   9
      Top             =   6600
      Width           =   11775
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打印上表(&U)"
         Height          =   375
         Left            =   4140
         TabIndex        =   16
         Top             =   165
         Width           =   1455
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "上表输出&Excel"
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   165
         Width           =   1455
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "设置(&Z)"
         Height          =   380
         Left            =   7530
         TabIndex        =   14
         Top             =   165
         Width           =   1250
      End
      Begin VB.CommandButton cmdAllSel 
         Caption         =   "全选(&A)"
         Height          =   380
         Left            =   105
         TabIndex        =   12
         ToolTipText     =   "快键:CTRL+A"
         Top             =   165
         Width           =   1250
      End
      Begin VB.CommandButton cmdALLCls 
         Caption         =   "全清(&S)"
         Height          =   380
         Left            =   1365
         TabIndex        =   11
         ToolTipText     =   "快键:CTRL+C"
         Top             =   165
         Width           =   1250
      End
      Begin VB.Frame fraBottomSplit 
         Height          =   30
         Left            =   -210
         TabIndex        =   10
         Top             =   0
         Width           =   12405
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "打印(&P)"
         Height          =   380
         Left            =   8865
         TabIndex        =   5
         Top             =   165
         Width           =   1250
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   380
         Left            =   10125
         TabIndex        =   6
         Top             =   165
         Width           =   1250
      End
   End
   Begin VB.PictureBox picSeach 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   11775
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   11775
      Begin VB.Frame fraSearch 
         Caption         =   "病区:一病区"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   840
         Left            =   90
         TabIndex        =   8
         Top             =   120
         Width           =   11235
         Begin VB.CommandButton cmd刷新 
            Caption         =   "刷新(&R)"
            Height          =   375
            Left            =   8310
            TabIndex        =   3
            Top             =   285
            Width           =   1125
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   1035
            TabIndex        =   1
            Top             =   330
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   183959555
            CurrentDate     =   36576
         End
         Begin VB.Label lbl截止日期 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "截止日期"
            Height          =   180
            Left            =   225
            TabIndex        =   0
            Top             =   390
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "通知单将打印病人在指定截止日期所在期间内的费用欠款情况！"
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   2775
            TabIndex        =   2
            Top             =   390
            Width           =   5040
         End
      End
      Begin VB.Image img16 
         Height          =   240
         Left            =   0
         Picture         =   "frmPatiPressMoney.frx":06EA
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPressMoney 
      Height          =   5430
      Left            =   75
      TabIndex        =   4
      Top             =   1080
      Width           =   11565
      _cx             =   20399
      _cy             =   9578
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPatiPressMoney.frx":0C74
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
      ExplorerBar     =   3
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
      Begin VB.PictureBox picImg 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   75
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   13
         Top             =   60
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmPatiPressMoney.frx":0CA1
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmPatiPressMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mlng病区 As Long, mblnOk As Boolean, mlng病人ID As Long
Private mlng主页ID As Long
Private mblnFirst As Boolean, mbytPrintModule As Byte
Private mlngPrintRow As Long    '当前正在打印的行
Private mstrPrintDate As Date
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1

Public Function zlPatiPressMoney(ByVal frmMain As Object, ByVal lngMoudle As Long, _
    ByVal strPrivs As String, ByVal lng病区 As Long, ByVal str病区名称 As String, _
    Optional lng病人ID As Long = 0, Optional bytPrintModule As Byte = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进入病人催款管理界面
    '入参:frmMain-调用的窗口
    '       bytPrintModule-2.打印;1-预览
    '出参:
    '返回:如果打印成功1个以上的病人,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-16 10:28:25
    '问题:35386
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    mbytPrintModule = bytPrintModule
    mlngModule = lngMoudle: mstrPrivs = strPrivs: mlng病区 = lng病区: mblnOk = False: mlng病人ID = lng病人ID
    If lng病区 = 0 And lng病人ID = 0 Then
        MsgBox "注意:" & vbCrLf & "    不支持对所有病区进行打印!", vbInformation + vbDefaultButton1 + vbOKOnly
        Exit Function
    End If
    If lng病人ID <> 0 Then
        '76451,冉俊明,2014-8-19
        gstrSQL = "Select 姓名,性别,年龄,主页ID From 病人信息 where 病人id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人信息", lng病人ID)
        If rsTemp.EOF Then '
            MsgBox "注意:" & vbCrLf & " 未找到相关的病人,不能继续!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        mlng主页ID = Nvl(rsTemp!主页ID)   '42626
        fraSearch.Caption = "姓名:" & rsTemp!姓名 & String(4, " ") & "性别:" & Nvl(rsTemp!性别) & String(4, " ") & "年龄:" & Nvl(rsTemp!年龄)
        If lng病区 <> 0 Then
            gstrSQL = "Select  Max(主页ID) as 主页ID From 病案主页 where 病人id=[1] And 当前病区ID=[2] "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人信息", lng病人ID, lng病区)
            If Val(Nvl(rsTemp!主页ID)) <> 0 Then mlng主页ID = Val(Nvl(rsTemp!主页ID))
        End If
    Else
        fraSearch.Caption = "『" & str病区名称 & "』的在院病人"
    End If
    
    mblnFirst = True
    Me.Show 1, frmMain
    zlPatiPressMoney = mblnOk
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdALLCls_Click()
    Dim i As Long
    With vsPressMoney
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("选择")) = 0
        Next
    End With
End Sub
Private Sub cmdAllSel_Click()
    Dim i As Long
    With vsPressMoney
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("选择")) = 1
            If Val(.TextMatrix(i, .ColIndex("催款金额"))) = 0 Then
                .TextMatrix(i, .ColIndex("催款金额")) = .Cell(flexcpData, i, .ColIndex("催款金额"))
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExcel_Click()
    Call PrintGrid(3)
End Sub

Private Sub cmdOK_Click()
    
    If zlPrintPatiPressMoney = False Then Exit Sub
    mblnOk = True
End Sub

Private Sub cmdPrint_Click()
    Call PrintGrid(1)
End Sub

Private Sub cmd刷新_Click()
    Call FillData
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mlng病人ID <> 0 Then cmdOK.SetFocus: Exit Sub
    vsPressMoney.SetFocus
    With vsPressMoney
         .Col = .ColIndex("选择")
        .Row = 1
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        Select Case KeyCode
        Case vbKeyA
            If Shift = vbCtrlMask Then cmdAllSel_Click
        Case vbKeyC
            If Shift = vbCtrlMask Then cmdALLCls_Click
        End Select
        If Not Me.ActiveControl Is vsPressMoney Then
            If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
        End If
End Sub

Private Sub Form_Load()
    RestoreWinState Me, Me.Name
    dtpEnd.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtpEnd.Value = DateAdd("d", -1, dtpEnd.MaxDate)
    Set mobjReport = New clsReport
    Call FillData
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsPressMoney
        .Left = Me.ScaleLeft + 50
        .Width = Me.ScaleWidth - 100
        .Top = picSeach.Top + picSeach.Height + 20
        .Height = Me.ScaleHeight - .Top - picDown.Height - 50
    End With
End Sub
 
Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    Dim i As Long
    i = mlngPrintRow
    With vsPressMoney
        Screen.MousePointer = 11
        If i < 1 Or i > .Rows - 1 Then Exit Sub
        '更新病人的缴款数据
         gstrSQL = "Zl_病案主页从表_首页整理("
        '    病人id_In 病案主页从表.病人id%Type,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("病人ID"))) & ","
        '    主页id_In 病案主页从表.主页id%Type,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("主页ID"))) & ","
        '    信息名_In 病案主页从表.信息名%Type,
        gstrSQL = gstrSQL & "'上次催款金额',"
        '    信息值_In 病案主页从表.信息值%Type
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("催款金额"))) & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        'ZL_病人催款记录_INSERT(
        gstrSQL = "ZL_病人催款记录_INSERT("
        '    病人ID_IN IN 病人催款记录.病人ID%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("病人ID"))) & ","
        '    主页ID_IN IN 病人催款记录.主页ID%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("主页ID"))) & ","
        '    预交余额_IN IN 病人催款记录.预交余额%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("预交余额"))) & ","
        '    未结费用_IN IN 病人催款记录.未结费用%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("未结费用"))) & ","
        '    自费金额_IN IN 病人催款记录.自费金额,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("未结费用"))) - Val(.TextMatrix(i, .ColIndex("医保预结"))) & ","
        '    医保预结_IN IN 病人催款记录.医保预结%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("医保预结"))) & ","
        '    当前余额_IN IN 病人催款记录.当前余额%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("预交余额"))) + Val(.TextMatrix(i, .ColIndex("医保预结"))) - Val(.TextMatrix(i, .ColIndex("未结费用"))) & ","
        '    催款下限_IN IN 病人催款记录.催款下限%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("催款下限"))) & ","
        '    催款标准_IN IN 病人催款记录.催款标准%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("催款标准"))) & ","
        '    催款金额_IN IN 病人催款记录.催款金额%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("催款金额"))) & ","
        '    打印日期_IN IN 病人催款记录.打印日期%TYPE,
        gstrSQL = gstrSQL & "to_date('" & mstrPrintDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '    打印人_IN IN 病人催款记录.打印人%TYPE
        gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        .Cell(flexcpData, i, .ColIndex("姓名")) = "1"
        .Cell(flexcpPicture, i, .ColIndex("姓名")) = img16.Picture
        .Cell(flexcpPictureAlignment, i, .ColIndex("姓名")) = 1
    End With
End Sub
 
Private Sub picSeach_Resize()
    Err = 0: On Error Resume Next
    With picSeach
        fraSearch.Left = .ScaleLeft + 50
        fraSearch.Top = .ScaleTop + 100
       ' fraSearch.Height = .ScaleHeight - 100
        fraSearch.Width = .ScaleWidth - 100
        cmd刷新.Left = .ScaleWidth - fraSearch.Left - cmd刷新.Width * 2
    End With
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        fraBottomSplit.Left = .ScaleLeft
        fraBottomSplit.Top = .ScaleTop
        fraBottomSplit.Width = .ScaleWidth
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - cmdCancel.Width / 2
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
        cmdSetup.Left = cmdOK.Left - cmdSetup.Width - 50
    End With
End Sub
Private Function FillData(Optional blnDefault As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:填充数据
    '入参:
    '出参:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-21 15:06:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strWhere As String, rsTemp As ADODB.Recordset, i As Long
    
    On Error GoTo errHandle
    strWhere = ""
    '当前在院的病人
    '42626
    If mlng病人ID = 0 Then strWhere = "   And A.病人ID=J1.病人ID And J1.病区ID=[1] And Nvl(B.状态,0)<>3 And A.主页ID=B.主页ID  "
     If mlng病区 > 0 And mlng病人ID <> 0 Then strWhere = strWhere & " And B.当前病区ID=[1]"
     If mlng病人ID > 0 Then strWhere = strWhere & " And B.病人ID=[2] And B.主页ID=[3]"
     '剩余款:Format((!预交余额 - !费用余额 + !预结费用+ Nvl(A.担保额, 0) ):Max(Nvl(A.担保额, 0)):37785
    '公式:预交金额+担保额+（医保病人报销总额）-未结费用余额-催款下限<0 就为本次欠款的病人
   strSql = "" & _
    "   Select A.病人ID, B.主页ID, B.状态, B.病人性质,B.出院科室id As 当前科室id, B.险类, B.当前病区ID,  " & _
    "            '1' as 选择,A.姓名, B.住院号, B.出院病床 As 床号, B.费别, A.性别, A.年龄, C.名称 As 当前科室, A.就诊卡号, E.密码,  " & _
    "           to_char(B.入院日期,'yyyy-mm-dd hh24:mi:ss') as 入院日期,  to_char(B.出院日期,'yyyy-mm-dd hh24:mi:ss') as 出院日期,  to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') as 登记时间, " & _
    "           ltrim(to_char(nvl(Max(M.催款下限),0),'9999999990.00')) As 催款下限, ltrim(to_char(nvl(Max(M.催款标准),0),'9999999990.00')) As 催款标准," & vbNewLine & _
    "           ltrim(to_char(Max(Nvl(X.预交余额, 0)),'9999999990.00')) As 预交余额,ltrim(to_char(Sum(Nvl(X1.金额, 0)),'9999999990.00')) As 医保预结, " & vbNewLine & _
    "           ltrim(to_char(Max(Nvl(A.担保额, 0)),'9999999990.00')) As 担保额,ltrim(to_char(Max(Nvl(X.费用余额, 0)),'9999999990.00')) As 未结费用," & vbNewLine & _
    "           ltrim(to_char(Max(Nvl(X.预交余额, 0))-Max(Nvl(X.费用余额, 0))+Sum(Nvl(X1.金额, 0)) ,'9999999990.00')) As 剩余款," & vbNewLine & _
    "           ltrim(to_char(case when to_number(Max(nvl(D1.信息值,'0')))> 0 then  to_number(Max(nvl(D1.信息值,'0')))  " & _
    "                               When  Max(nvl(X.预交余额,0))+Max(Nvl(A.担保额, 0))+Sum(nvl(X1.金额,0))-Max(nvl(x.费用余额,0))<0 then round(abs(Max(nvl(X.预交余额,0))+Max(Nvl(A.担保额, 0))+Sum(nvl(X1.金额,0))-Max(nvl(x.费用余额,0)))/100,0)*100+Max(nvl(M.催款标准,0)) " & _
    "                               Else   Max(M.催款标准) end,'9999999990.00')) As 催款金额," & vbNewLine & _
    "           Nvl(E.医保号, D.信息值) 医保号, A.家庭电话, B.医疗付款方式, B.审核人, B.病人类型, H.名称 当前病区" & vbNewLine & _
    "   From 病人信息 A, 病案主页 B, 病案主页从表 D, 病案主页从表 D1, 医保病人档案 E, 医保病人关联表 F, 病人余额 X,保险模拟结算 X1, " & vbNewLine & _
    "         记帐报警线 M,部门表 C, 部门表 H" & IIf(mlng病人ID = 0, ",在院病人 J1", "") & vbNewLine & _
    "   Where A.病人ID = B.病人ID And B.出院科室ID = C.ID And Nvl(B.主页ID, 0) <> 0 " & vbNewLine & _
    "          And B.病人ID = D.病人ID(+) And B.主页ID = D.主页ID(+) And  D.信息名(+) = '医保号'  " & vbNewLine & _
    "          And B.病人ID = D1.病人ID(+) And B.主页ID = D1.主页ID(+) And  D1.信息名(+) = '上次催款金额'  " & vbNewLine & _
    "          And A.病人ID = X.病人ID(+) And X.性质(+) = 1 And X.类型(+)=2 And B.病人ID=X1.病人ID(+) and B.主页id=X1.主页ID(+)  " & vbNewLine & _
    "          And A.病人ID = F.病人ID(+) And F.标志(+) = 1 And F.医保号 = E.医保号(+) And F.险类 = E.险类(+) And F.中心 = E.中心(+)  " & vbNewLine & _
    "          And B.当前病区ID = H.ID And (H.站点='" & gstrNodeNo & "' Or H.站点 is Null)" & vbNewLine & strWhere & vbNewLine & _
    "          And B.当前病区ID=M.病区ID(+) And zl_PatiWarnScheme(b.病人id,b.主页ID) =M.适用病人(+) " & vbNewLine & _
    "   Group by A.病人ID, B.主页ID, A.登记时间, B.状态, B.病人性质, B.数据转出, B.出院科室id, A.就诊卡号, B.险类, E.密码, B.当前病区ID, A.姓名, B.住院号, B.出院病床,B.费别, A.性别, A.年龄, B.入院日期, B.出院日期, C.名称, Decode(Nvl(X.费用余额, 0), 0, '√', ''), " & vbNewLine & _
    "          Nvl(E.医保号, D.信息值), A.家庭电话, B.医疗付款方式, B.审核人, B.病人类型, H.名称 " & _
         IIf(mlng病人ID <> 0, "", "   having (Max(nvl(X.预交余额,0))+Max(Nvl(A.担保额, 0))+Sum(nvl(X1.金额,0))-Max(nvl(x.费用余额,0))-Max(nvl(M.催款下限,0)))<0 " & vbNewLine) & _
         IIf(mlng病区 = 0, " Order by 住院号 Desc", " Order by LPAD(床号,10,' ')")
 
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病区, mlng病人ID, mlng主页ID)
    With vsPressMoney
        .Clear 0: .Cols = 1
        .FixedCols = 1
       Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        For i = 1 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            .ColData(i) = "0||1"
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "状态" Or .ColKey(i) = "密码" Or .ColKey(i) = "病人性质" Or .ColKey(i) = "险类" Then
                .ColHidden(i) = True
                ' ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                .ColData(i) = "-1||1"
            End If
            If .ColKey(i) Like "*标准*" Or .ColKey(i) Like "*下限*" Or .ColKey(i) Like "*额*" Then
                 .ColAlignment(i) = flexAlignRightCenter
            End If
            '屏蔽以下列宽
            Select Case .ColKey(i)
            Case "住院号", "床号", "性别", "年龄", "当前科室", "病人类型"
            Case "预交余额", "未结费用", "催款下限", "催款标准", "剩余款", "医保预结"
                 .ColAlignment(i) = flexAlignRightCenter
            Case "姓名", "催款金额", "选择"
                   .ColData(i) = "1||0"
                   If .ColKey(i) = "选择" Then .ColDataType(i) = flexDTBoolean
                   If .ColKey(i) = "选择" Then .ColAlignment(i) = flexAlignCenterCenter
            Case Else
                .ColHidden(i) = True
            End Select
        Next
        '设置颜色
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("险类"))) <> 0 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            End If
            If Val(.TextMatrix(i, .ColIndex("预交余额"))) = 0 Then .TextMatrix(i, .ColIndex("预交余额")) = ""
            If Val(.TextMatrix(i, .ColIndex("未结费用"))) = 0 Then .TextMatrix(i, .ColIndex("未结费用")) = ""
            If Val(.TextMatrix(i, .ColIndex("催款下限"))) = 0 Then .TextMatrix(i, .ColIndex("催款下限")) = ""
            If Val(.TextMatrix(i, .ColIndex("催款标准"))) = 0 Then .TextMatrix(i, .ColIndex("催款标准")) = ""
            If Val(.TextMatrix(i, .ColIndex("剩余款"))) = 0 Then .TextMatrix(i, .ColIndex("剩余款")) = ""
            If Val(.TextMatrix(i, .ColIndex("医保预结"))) = 0 Then .TextMatrix(i, .ColIndex("医保预结")) = ""
            
        Next
        '自动列宽
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsPressMoney, Me.Caption, "催款列表", False
        If .ColIndex("标志") >= 0 Then .ColWidth(.ColIndex("标志")) = 300
        .Cell(flexcpBackColor, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = &HE7CFBA
        .Cell(flexcpBackColor, 1, .ColIndex("催款金额"), .Rows - 1, .ColIndex("催款金额")) = &HE7CFBA
        .Redraw = flexRDBuffered
    End With
    
    FillData = True
    Exit Function
errHandle:
    vsPressMoney.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
    SaveWinState Me, Me.Name
    zl_vsGrid_Para_Save mlngModule, vsPressMoney, Me.Caption, "催款列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub
 
Private Sub vsPressMoney_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        With vsPressMoney
            Select Case Col
            Case .ColIndex("催款金额")
                .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, .Col)), "0.00")
                If Val(.TextMatrix(Row, Col)) <> 0 And GetVsGridBoolColVal(vsPressMoney, Row, .ColIndex("选择")) = False Then
                    vsPressMoney.TextMatrix(Row, .ColIndex("选择")) = 1
                ElseIf Val(.TextMatrix(Row, Col)) = 0 Then
                    vsPressMoney.TextMatrix(Row, .ColIndex("选择")) = 0
                End If
            Case .ColIndex("选择")
                If GetVsGridBoolColVal(vsPressMoney, Row, Col) Then
                    If Val(.TextMatrix(Row, .ColIndex("催款金额"))) = 0 Then
                        .TextMatrix(Row, .ColIndex("催款金额")) = Format(Val(.Cell(flexcpData, Row, .ColIndex("催款金额"))), "0.00")
                    End If
                End If
            Case Else
            End Select
        End With
End Sub
Private Sub vsPressMoney_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsPressMoney, Me.Caption, "催款列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsPressMoney_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsPressMoney, Me.Caption, "催款列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlcontrol.GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsPressMoney, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsPressMoney, Me.Caption, "催款列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub
Private Sub picImg_Click()
    Call imgCol_Click
End Sub
Private Function zlPrintPatiPressMoney() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印病人缴款通知单
    '编制:刘兴洪
    '日期:2010-12-16 15:15:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim i As Long, blnData As Boolean
    Dim str截止日期 As String
    Dim lngCount As Long
    '先检查可记帐数量
    With vsPressMoney
        blnData = False
        mstrPrintDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsPressMoney, i, .ColIndex("选择")) Then
                If Val(.TextMatrix(i, .ColIndex("催款金额"))) <= 0 Then
                    MsgBox "注意:" & "    在第" & i & "行中的催款金额必须大于零,请检查!", vbOKOnly + vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("催款金额")
                    If .RowIsVisible(.Row) = False Or .ColIsVisible(.Col) = False Then
                        Call .ShowCell(.Row, .Col)
                    End If
                    Exit Function
                End If
                lngCount = lngCount + 1
                blnData = True
            End If
        Next
    End With
    
    If blnData = False Then
        MsgBox "注意:" & "    未选择指定的打印数据,请检查!", vbOKOnly + vbInformation, gstrSysName
        vsPressMoney.SetFocus
        Exit Function
    End If
    If MsgBox("你是否真要打印这" & lngCount & "个病人的催款单吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    str截止日期 = Format(dtpEnd.Value, "yyyy-mm-dd")
    
    With vsPressMoney
        Screen.MousePointer = 11
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsPressMoney, i, .ColIndex("选择")) And Val(.TextMatrix(i, .ColIndex("病人ID"))) > 0 Then
                mlngPrintRow = i
                Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1139_3", Me, "病人ID=" & Val(.TextMatrix(i, .ColIndex("病人ID"))), _
                    "日期=" & str截止日期, "催款金额=" & Val(.TextMatrix(i, .ColIndex("催款金额"))), mbytPrintModule)
            Else
               .Cell(flexcpData, i, .ColIndex("姓名")) = "0"
            End If
        Next
    End With
    Screen.MousePointer = 0
    MsgBox "所有病人打印完成！", vbInformation, gstrSysName
    zlPrintPatiPressMoney = True
End Function
Private Sub vsPressMoney_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPressMoney
        Select Case Col
        Case .ColIndex("催款金额"), .ColIndex("选择")
            Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsPressMoney_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPressMoney
        Select Case Col
        Case .ColIndex("标志")
            Cancel = True
        Case Else
        End Select
    End With
End Sub

Private Sub vsPressMoney_EnterCell()
    '暂未设置
    With vsPressMoney
    End With
End Sub

Private Sub vsPressMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPressMoney
        If Val(.TextMatrix(.Row, .ColIndex("病人ID"))) = 0 Or (.Col >= .ColIndex("催款金额") And .Row = .Rows - 1) Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    With vsPressMoney
        Select Case .Col
        Case .ColIndex("催款金额")
                If .Row < .Rows - 1 Then
                    .Col = .Col: .Row = .Row + 1
                End If
        Case .ColIndex("选择")
                If .ColIndex("选择") > .ColIndex("催款金额") Then
                   .Col = .ColIndex("催款金额"): .Row = .Row + 1
                Else
                    .Col = .ColIndex("催款金额")
                End If
        Case Else
        End Select
    End With
        
    End With
End Sub

Private Sub vsPressMoney_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '编辑处理
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPressMoney
        Select Case Col
        Case .ColIndex("催款金额")
                If Row < .Rows - 1 Then
                    .Col = Col: .Row = .Row + 1
                End If
        Case .ColIndex("选择")
                If .ColIndex("选择") > .ColIndex("催款金额") Then
                   .Col = .ColIndex("催款金额"): .Row = .Row + 1
                Else
                    .Col = .ColIndex("催款金额")
                End If
        Case Else
        End Select
    End With
End Sub

Private Sub vsPressMoney_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsPressMoney_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsPressMoney
        Select Case .Col
            Case .ColIndex("催款金额")
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    If KeyAscii = vbKeyBack Then Exit Sub
                    If KeyAscii = vbKeyReturn Then Exit Sub
                    If KeyAscii = Asc(".") Then
                        If InStr(1, .EditText, ".") = 0 Then
                            Exit Sub
                        End If
                    End If
                    KeyAscii = 0
                End If
            Case Else
            
        End Select
    End With
End Sub

Private Sub vsPressMoney_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    '数据验证
    With vsPressMoney
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
            Case .ColIndex("催款金额")
                If Val(strKey) > 999999999 Then
                    MsgBox "注意:" & vbCrLf & "    催款金额只能在0-999999999范围中!"
                    Cancel = True
                End If
                strKey = Format(Val(strKey), "0.00")
                .EditText = strKey
                .TextMatrix(Row, .Col) = strKey
        End Select
    End With
End Sub
  
Private Sub cmdSetup_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1139_3", Me
End Sub

Private Sub PrintGrid(bytMode As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印表格内容
    '入参:bytMode:1-打印,2-预览,3-输出到Excel
    '编制:刘兴洪
    '日期:2011-05-13 10:10:23
    '问题:37934
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPrintObject  As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim lngCol As Long, lngRow As Long, i As Long
    Dim cllCol As New Collection

    Err = 0: On Error GoTo errHandle
    '记录列状态
    With vsPressMoney
        .Redraw = flexRDNone
        lngRow = .Row: lngCol = .Col
        For i = 0 To .Cols - 1
             cllCol.Add Array(CStr(.ColData(i)), .ColWidth(i), IIf(.ColHidden(i), 1, 0)), "K" & i
             If i = .ColIndex("标志") Then .ColWidth(i) = 0
             If .ColHidden(i) Then .ColWidth(i) = 0
        Next
    End With
        
    '表头
    objPrintObject.Title.Text = "病人催款表"
    objPrintObject.Title.Font.Name = "楷体_GB2312"
    objPrintObject.Title.Font.Size = 18
    objPrintObject.Title.Font.Bold = True
    '表项
    objRow.Add fraSearch.Caption
    objRow.Add "截止日期：" & Format(dtpEnd.Value, "yyyy年mm月DD日")
    objPrintObject.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日 HH:MM:SS")
    objPrintObject.BelowAppRows.Add objRow
    '表体
    Set objPrintObject.Body = vsPressMoney
    
    '输出
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrintObject)
        Me.Refresh
        If bytMode <> 0 Then zlPrintOrView1Grd objPrintObject, bytMode
    Else
        zlPrintOrView1Grd objPrintObject, bytMode
    End If
    '恢复原始状态
     With vsPressMoney
         .Row = lngRow: .Col = lngCol
        For i = 1 To cllCol.Count
             .ColData(i - 1) = cllCol(i)(0)
             .ColWidth(i - 1) = cllCol(i)(1)
             .ColHidden(i - 1) = IIf(cllCol(i)(2) = 1, True, False)
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    '恢复原始状态
     With vsPressMoney
         .Row = lngRow: .Col = lngCol
        For i = 1 To cllCol.Count
             .ColData(i - 1) = cllCol(i)(0)
             .ColWidth(i - 1) = cllCol(i)(1)
             .ColHidden(i - 1) = IIf(cllCol(i)(2) = 1, True, False)
        Next
        .Redraw = flexRDBuffered
    End With
End Sub
