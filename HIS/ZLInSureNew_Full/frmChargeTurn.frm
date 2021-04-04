VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmChargeTurn 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "门(急)诊费用转住院"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   8715
   ControlBox      =   0   'False
   Icon            =   "frmChargeTurn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   90
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   8400
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3180
      Width           =   8400
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8715
      TabIndex        =   12
      Top             =   0
      Width           =   8715
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   6000
         TabIndex        =   2
         Top             =   95
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   3720
         TabIndex        =   1
         Top             =   90
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   609
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   225574915
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   90
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   90963971
         CurrentDate     =   36588
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   180
         Left            =   3480
         TabIndex        =   14
         Top             =   180
         Width           =   90
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费发生时间"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   1080
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1905
      Left            =   30
      TabIndex        =   4
      ToolTipText     =   "双击单据查看明细"
      Top             =   3240
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   3360
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmChargeTurn.frx":058A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   8715
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5280
      Width           =   8715
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   60
         TabIndex        =   8
         Top             =   45
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全清(&C)"
         Height          =   350
         Index           =   1
         Left            =   2640
         TabIndex        =   6
         Top             =   45
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全选(&A)"
         Height          =   350
         Index           =   0
         Left            =   1455
         TabIndex        =   5
         Top             =   45
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "退出(&X)"
         Height          =   350
         Left            =   7380
         TabIndex        =   7
         Top             =   45
         Width           =   1100
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2535
      Left            =   30
      TabIndex        =   3
      ToolTipText     =   "双击单据查看明细"
      Top             =   600
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   4471
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmChargeTurn.frx":08A4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   5760
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeTurn.frx":0BBE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
End
Attribute VB_Name = "frmChargeTurn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)

Private mstrNOS As String
Private mlngPatient As Long

Private Enum 交易Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
End Enum

Private Enum 医院业务
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
    Support多单据收费必须全退 = 39  '多单据收费必须全退
End Enum

Private Enum COL
    C0选择 = 0
    C1类别 = 1
    C2医保 = 2
    C3单据号 = 3
    C4票据号 = 4
    C5开单人 = 5
    C6应收金额 = 6
    C7实收金额 = 7
    C8发生时间 = 8
    C9结帐ID = 9
    C10险类 = 10
End Enum
Private mbln门诊转住院先审核 As Boolean


Public Sub ShowME(objParent As Object, ByVal lngPatient As Long, ByRef strNOS As String)
'参数:lngPatient-病人ID
'返回:要进行费用转入的单据,票据,结帐ID,险类(非医保为零):H0000001,F000023,81235,901;H0000002,F000045,81263,901;...
    mlngPatient = lngPatient
    mstrNOS = strNOS
    
    '此时会先隐式调用事件Form_Load
    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
    Call SetBillSelected(strNOS)
    
    Call Me.Show(vbModal, objParent)
    
    strNOS = mstrNOS
End Sub

Private Sub SetBillSelected(ByVal strNOS As String)
'说明:如果转入几天后失败,再进入选择窗体,以前选择的且已被转入的单据现在是"不可转入",所以不应被选择
    Dim i As Long
    With mshList
        For i = 1 To .Rows - 1
            If InStr(";" & strNOS, ";" & .TextMatrix(i, COL.C3单据号)) > 0 And .TextMatrix(i, COL.C1类别) = "可转入" Then
                .TextMatrix(i, COL.C0选择) = "√"
            Else
                .TextMatrix(i, COL.C0选择) = ""
            End If
        Next
    End With
End Sub

'功能:检查入院时间之后是否存在转入数据
'返回:转入数据的登记时间
Public Function CheckExistTurn(ByVal lngPatient As Long, ByRef dat入院时间 As Date) As Boolean
    Dim rsTmp As New ADODB.Recordset, strSQL As String
        
    On Error GoTo ErrH
    strSQL = "Select Max(发生时间) 发生时间 From 住院费用记录" & vbNewLine & _
            "Where 记录性质 = 2 And 记录状态 In(1,3) And 病人id = [1] And 主页id Is Null And 标识号 Is Null"

    'His9低版本公共部件不支持绑定变量方式
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查是否存在已转单据", lngPatient)
    If ChkRsState(rsTmp) Then Exit Function
    If Not IsNull(rsTmp!发生时间) Then
        dat入院时间 = rsTmp!发生时间
        CheckExistTurn = True
    End If
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'功能:更新记帐单的主页ID
Public Sub ExecteUpdate(ByVal lngPatient As Long, ByVal strInID As String, ByVal lngPageID As Long, ByVal dat入院时间 As Date)
    Dim strSQL As String
    
    On Error GoTo ErrH
    strSQL = "Zl_门诊费用转住院_Update(" & lngPatient & "," & strInID & "," & lngPageID & _
            ",To_Date('" & Format(dat入院时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    Call zlDatabase.ExecuteProcedure(strSQL, "更新记帐单")
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'功能:根据指定的单据号序列,执行门诊费用转住院费用,及医保退费结算操作
'参数:strNOS-要进行费用转入的单据,票据,结帐ID,险类(非医保为零):H0000001,F000023,81235,901;H0000002,F000045,81263,901;...
'     strInID-住院号,lngPageID-主页ID,这两个参数仅在医保入院补充登记时才传入
Public Function ExecuteTurn(ByVal strNOS As String, ByVal strInID As String, ByVal lngPageID As Long, ByVal dat入院时间 As Date, ByVal lng入院科室ID As Long) As Boolean

    Dim DateDel As Date, arrNO As Variant, arrInfo As Variant
    Dim i As Long, j As Long, lngcnt As Long
    Dim strSQL As String, strInvoice As String, strInDate As String, strDelDate As String
    
    Dim blnTrans As Boolean, blnTransMedicare As Boolean, blnDo As Boolean
    Dim intinsure As Integer, strAdvance As String
    
    If strNOS = "" Then Exit Function
    
    strInDate = "To_Date('" & Format(dat入院时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    strDelDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    arrNO = Split(strNOS, ";")
    
    On Error GoTo ErrH
    i = LBound(arrNO)
    Do While i <= UBound(arrNO)
        lngcnt = 1
        strInvoice = Trim(Split(arrNO(i), ",")(1))
        If strInvoice <> "" Then
            For j = i + 1 To UBound(arrNO)
                If strInvoice = Split(arrNO(j), ",")(1) Then
                    lngcnt = lngcnt + 1
                Else
                    Exit For
                End If
            Next
        End If
        
        '医保要求从最后一张开始退
        For j = i To i + lngcnt - 1
            gcnOracle.BeginTrans: blnTrans = True
            arrInfo = Split(arrNO(j), ",")
            
            strSQL = "Zl_门诊费用转住院_insert('" & arrInfo(0) & "'," & IIf(Val(strInID) = 0, "Null", strInID) & _
                "," & IIf(lngPageID = 0, "Null", lngPageID) & "," & strInDate & "," & lng入院科室ID & "," & _
                strDelDate & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "门诊费用转住院")
            
            blnTransMedicare = False
            intinsure = Val(arrInfo(3))
            If intinsure <> 0 Then
                strAdvance = lngcnt & "|" & (j - i + 1)
                
                '$IF HIS9.19
                #If gverControl = 0 Then
                    blnDo = gclsInsure.ClinicDelSwap(Val(arrInfo(2)))
                #ElseIf gverControl = 1 Then
                '$ELSE  HIS+
                    blnDo = gclsInsure.ClinicDelSwap(Val(arrInfo(2)), , intinsure)
                #Else
                    blnDo = gclsInsure.ClinicDelSwap(Val(arrInfo(2)), , intinsure, strAdvance)
                #End If
                '$END IF
                
                If Not blnDo Then
                    GoTo ErrH
                Else
                    blnTransMedicare = True
                End If
            End If
            gcnOracle.CommitTrans: blnTrans = False
            
            #If gverControl >= 2 Then
                If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intinsure)
            #End If
        Next
               
        i = i + lngcnt
    Loop

    ExecuteTurn = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        gcnOracle.RollbackTrans
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        #If gverControl >= 2 Then
            If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, intinsure)
        #End If
    End If
    
    Call SaveErrLog
End Function


Private Sub ShowBills(ByVal lngPatient As Long, ByVal datBegin As Date, ByVal datEnd As Date)
'功能:读取并显示病人指定天数内的门诊费用单据
    Dim i As Long, DatTmp As Date, strSQL As String
    Dim rsList As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    
    If datBegin > datEnd Then
        DatTmp = datEnd
        datEnd = datBegin
        datBegin = DatTmp
    End If
    strBegin = "To_Date('" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    strEnd = "To_Date('" & Format(datEnd, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    sta.Panels(2).Text = "正在读取收费单据,请稍候 ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    
    On Error GoTo ErrH
    strSQL = "Select '√' as 选择,'可转入' as 类别,Decode(B.险类,Null,'','√') as 医保, A.NO As 单据号, A.实际票号 As 票据号, A.开单人," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.应收金额), '900090009" & gstrDec & "')) As 应收金额," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.实收金额), '900090009" & gstrDec & "')) As 实收金额," & vbNewLine & _
            "       To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间, A.结帐ID, B.险类" & vbNewLine & _
            "From 门诊费用记录 A,保险结算记录 B" & vbNewLine & _
            "Where A.记录性质 = 1 And A.记录状态 = 1 And A.病人id+0 = [1] And A.发生时间 Between [2] And [3] And A.结帐ID = B.记录ID(+) And B.性质(+) = 1" & vbNewLine & _
            IIf(mbln门诊转住院先审核, "           And Exists(Select 1 From 门诊费用记录 M,费用审核记录 J where A.ID=J.费用ID and M.NO=A.NO And M.记录性质=A.记录性质 And J.审核日期 is Not NULL and  nvl(J.记录状态,0)=0 and J.性质=1) " & vbNewLine, "") & _
            "Group By A.NO, A.实际票号, A.开单人, A.发生时间, A.结帐ID, B.险类" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select '' as 选择,'不可转入' as 类别,Decode(B.险类,Null,'','√') as 医保, A.NO As 单据号, A.实际票号 As 票据号, A.开单人," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.应收金额), '900090009" & gstrDec & "')) As 应收金额," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.实收金额), '900090009" & gstrDec & "')) As 实收金额," & vbNewLine & _
            "       To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间,A.结帐ID,0 as 险类" & vbNewLine & _
            "From 门诊费用记录 A,保险结算记录 B" & vbNewLine & _
            "Where Mod(A.记录性质,10)=1 And A.记录状态 = 3 And A.病人id+0 = [1] And A.发生时间 Between [2] And [3] And A.结帐ID = B.记录ID(+) And B.性质(+) = 1" & vbNewLine & _
            "Group By A.NO, A.实际票号, A.开单人, A.发生时间, A.结帐ID, B.险类" & vbNewLine & _
            "Order By 类别, 票据号, 单据号 Desc"

    '注意:由于多单据退费要从最后一张开始退,所以排序很关键
    Set rsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatient, datBegin, datEnd)
    mshList.Redraw = False
    mshList.ClearStructure
    mshList.Clear
    mshList.Rows = 2
    
    If rsList.EOF Then
        sta.Panels(2).Text = "没有找到指定时间范围的收费单据!"
    Else
        Set mshList.DataSource = rsList
        sta.Panels(2).Text = "共 " & rsList.RecordCount & " 张收费单据"
    End If
    Call SetHeader
    Call SetBillColor
    mshList.Redraw = True
    
    mshList.Row = 1
    mshList.COL = 0: mshList.ColSel = mshList.Cols - 1
    Call mshList_EnterCell
    Screen.MousePointer = 0
    Me.Refresh
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "选择,4,500|类别,4,850|医保,4,500|单据号,4,850|票据号,4,1100|开单人,4,800|应收金额,7,850|实收金额,7,850|发生时间,4,1850|结帐ID,4,0|险类,4,0"
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        .COL = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Sub SetBillColor()
    Dim i As Long, j As Long
    
    With mshList
        For i = 1 To .Rows - 1
            .Row = i
            For j = 0 To .Cols - 1
                .COL = j
                If .TextMatrix(i, COL.C1类别) = "不可转入" Then
                    .CellForeColor = &H8000000C
                Else
                    .CellForeColor = 0
                End If
            Next
        Next
    End With
End Sub

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim strTmp As String, Datsys As Date
    
    Call RestoreWinState(Me, App.ProductName)
    mbln门诊转住院先审核 = IIf(Val(GetPara("门诊转住院先审核", glngSys, 1143, 0)) = 1, True, False)
        
    Datsys = zlDatabase.Currentdate
    strTmp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "开始时间")
    If IsDate(strTmp) Then
        dtpBegin.Value = CDate(strTmp)
    Else
        dtpBegin.Value = DateAdd("d", -3, Datsys)
    End If
        
    If mstrNOS <> "" Then
        strTmp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "结束时间")
    Else
        strTmp = ""
    End If
    If IsDate(strTmp) Then
        dtpEnd.Value = CDate(strTmp)
    Else
        dtpEnd.Value = Datsys
    End If
        
    Call SetHeader
    Call SetDetail
End Sub

Private Sub cmdExit_Click()
    Dim i As Long
    
    mstrNOS = ""
    With mshList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, COL.C1类别) = "可转入" And .TextMatrix(i, COL.C0选择) = "√" Then
                mstrNOS = mstrNOS & ";" & .TextMatrix(i, COL.C3单据号) & "," & .TextMatrix(i, COL.C4票据号) & _
                        "," & .TextMatrix(i, COL.C9结帐ID) & "," & .TextMatrix(i, COL.C10险类)
            End If
        Next
    End With
    mstrNOS = Mid(mstrNOS, 2)
    
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdRefresh_Click()
    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
    If cmdAll(0).Visible And cmdAll(0).Enabled Then Call cmdAll(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = asc("'") Then KeyAscii = 0
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    mshList.Width = Me.ScaleWidth - 100
    mshDetail.Width = Me.ScaleWidth - 100
    cmdExit.Left = picBottom.Left + picBottom.Width - cmdExit.Width - 100
    
    pic.Top = picBottom.Top - mshDetail.Height - 100
    mshDetail.Top = pic.Top + 50
    mshList.Height = pic.Top - mshList.Top - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "开始时间", Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss")
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "结束时间", Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mshList.TextMatrix(mshList.Row, COL.C3单据号) = "" Then Exit Sub
    Call SetRowSelected(mshList.Row, Trim(mshList.TextMatrix(mshList.Row, COL.C0选择)) = "")
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
    If mshList.TextMatrix(mshList.Row, COL.C3单据号) = "" Then Exit Sub
    If KeyAscii = 32 Then Call SetRowSelected(mshList.Row, Trim(mshList.TextMatrix(mshList.Row, COL.C0选择)) = "")
End Sub


Private Sub cmdAll_Click(Index As Integer)
    Dim i As Long
    
    With mshList
        .Redraw = False
        For i = 1 To .Rows - 1
            If Not SetRowSelected(i, Index = 0) Then
                .Row = i: .COL = 0: .ColSel = .Cols - 1
                Call mshList_EnterCell
                Exit For
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Function SetRowSelected(ByVal lngRow As Long, blnSelect As Boolean) As Boolean
'功能:设置一行的选择状态
'     如果是多张单据中的一张,则还需同时设置多张中的其它单据
    Dim intinsure As Integer, strNO As String, i As Long, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant, blnDo As Boolean
    
    With mshList
        If .TextMatrix(lngRow, COL.C1类别) = "可转入" And .TextMatrix(lngRow, COL.C0选择) <> IIf(blnSelect, "√", "") Then
            intinsure = Val(.TextMatrix(lngRow, COL.C10险类))
            
            If intinsure > 0 And blnSelect Then
                strNO = .TextMatrix(lngRow, COL.C3单据号)
                #If gverControl = 0 Then
                    blnDo = gclsInsure.GetCapability(support门诊结算作废)
                #Else
                    blnDo = gclsInsure.GetCapability(support门诊结算作废, , intinsure)
                #End If
                
                If Not blnDo Then
                    sta.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持门诊结算作废,此行不允许选择转入!"
                    .TextMatrix(lngRow, COL.C0选择) = ""
                    Exit Function
                Else
                    '再判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
                    strTmp = GetBalanceType(strNO)
                    If strTmp <> "" Then
                        arrBalanceType = Split(strTmp, ",")
                        For i = 0 To UBound(arrBalanceType)
                            strBalanceType = arrBalanceType(i)
                            blnDo = True
                            #If gverControl >= 2 Then
                                blnDo = gclsInsure.GetCapability(support门诊结算作废, , intinsure, strBalanceType)
                            #End If
                            
                            If Not blnDo Then
                                sta.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持" & strBalanceType & "作废,此行不允许选择转入!"
                                .TextMatrix(lngRow, COL.C0选择) = ""
                                Exit Function
                            End If
                        Next
                    End If
                End If
            End If
            
            .TextMatrix(lngRow, COL.C0选择) = IIf(blnSelect, "√", "")
            If intinsure > 0 Then   '全部选择或取消
                #If gverControl = 0 Then
                    blnDo = gclsInsure.GetCapability(Support多单据收费必须全退)
                #Else
                    blnDo = gclsInsure.GetCapability(Support多单据收费必须全退, , intinsure)
                #End If
                
                If blnDo Then
                    If Not SetMultiOther(lngRow, blnSelect, intinsure) Then Exit Function
                End If
            End If
        End If
    End With
    SetRowSelected = True
End Function

Private Function SetMultiOther(ByVal lngRow As Long, blnSelect As Boolean, intinsure As Integer) As Boolean
'功能:多张单据整体选择或取消
'     如果医保多张单据要求整体退费,选择其中一张时,全选多张,取消时全取消
    Dim i As Long, j As Long, k As Long, strNO As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant, blnDo As Boolean
    
    With mshList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, COL.C1类别) = "可转入" And .TextMatrix(i, COL.C4票据号) = .TextMatrix(lngRow, COL.C4票据号) And i <> lngRow Then
                If .TextMatrix(i, COL.C0选择) <> .TextMatrix(lngRow, COL.C0选择) Then
                   If intinsure <> 0 And blnSelect Then
                        strNO = .TextMatrix(i, COL.C3单据号)
                        '判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
                         strTmp = GetBalanceType(strNO)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                strBalanceType = arrBalanceType(j)
                                 
                                blnDo = True
                                #If gverControl >= 2 Then
                                    blnDo = gclsInsure.GetCapability(support门诊结算作废, , intinsure, strBalanceType)
                                #End If
                                 
                                If Not blnDo Then
                                    sta.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持" & strBalanceType & "作废,此行不允许选择转入!"
                                    For k = 1 To .Rows - 1
                                       If .TextMatrix(k, COL.C4票据号) = .TextMatrix(i, COL.C4票据号) Then
                                           .TextMatrix(k, COL.C0选择) = ""
                                       End If
                                    Next
                                    Exit Function
                                End If
                             Next
                         End If
                    End If
                    .TextMatrix(i, COL.C0选择) = IIf(blnSelect, "√", "")
                End If
            End If
        Next
    End With
    SetMultiOther = True
End Function

Private Function GetBalanceType(ByVal strNO As String) As String
'功能:获取一张单据中的医保结算方式串
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim i As Long
        
    On Error GoTo ErrH
    strSQL = "Select A.结算方式 From 病人预交记录 A, 结算方式 B" & vbNewLine & _
            "Where A.结算方式 = B.名称 And B.性质 In (3, 4) And A.NO =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    For i = 1 To rsTmp.RecordCount
        GetBalanceType = GetBalanceType & "," & rsTmp!结算方式
        rsTmp.MoveNext
    Next
    GetBalanceType = Mid(GetBalanceType, 2)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mshList_EnterCell()
    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, COL.C3单据号) = "" Then
        mshDetail.Clear
        mshDetail.Rows = 2
        Call SetDetail
    Else
        Call ShowDetail(mshList.TextMatrix(mshList.Row, COL.C3单据号))
    End If
    
    If mshList.TextMatrix(mshList.Row, COL.C1类别) = "不可转入" Then
        mshList.ForeColorSel = mshList.CellForeColor
    Else
        mshList.ForeColorSel = &H80000005
    End If
End Sub


Private Sub ShowDetail(ByVal strNO As String)
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo ErrH
    strSQL = "Select C.名称 As 类别, Nvl(E.名称, B.名称) As 名称, B.规格, A.计算单位 As 单位, Avg(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.标准单价), '900090.00000')) As 单价, LTrim(To_Char(Sum(A.应收金额), '90009" & gstrDec & "')) As 应收金额," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.实收金额), '90009" & gstrDec & "')) As 实收金额, D.名称 As 执行科室" & vbNewLine & _
            "From 门诊费用记录 A, 收费项目目录 B, 收费项目类别 C, 部门表 D, 收费项目别名 E" & vbNewLine & _
            "Where A.收费细目id = B.ID And A.收费类别 = C.编码 And A.执行部门id = D.ID(+) And A.NO = '" & strNO & "' And A.记录性质 = 1 And" & vbNewLine & _
            "      A.记录状态 In (1, 3) And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = 3" & vbNewLine & _
            "Group By Nvl(A.价格父号, A.序号), C.名称, Nvl(E.名称, B.名称), B.规格, A.计算单位, D.名称" & vbNewLine & _
            "Order By Nvl(A.价格父号, A.序号)"
    'Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "显示明细")
    
    mshDetail.Redraw = False
    mshDetail.ClearStructure
    mshDetail.Clear
    mshDetail.Rows = 2
    If Not rsTmp.EOF Then Set mshDetail.DataSource = rsTmp
    Call SetDetail
    mshDetail.Redraw = True
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "类别,1,650|名称,1,1500|规格,1,1450|单位,4,500|数量,7,500|单价,7,850|应收金额,7,850|实收金额,7,850|执行科室,4,1000"
    With mshDetail
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        
        .COL = 0: .ColSel = .Cols - 1
    End With
End Sub


Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If mshList.Height + y < 600 Or mshDetail.Height - y < 800 Then Exit Sub
        pic.Top = pic.Top + y
        mshList.Height = mshList.Height + y
        mshDetail.Top = mshDetail.Top + y
        mshDetail.Height = mshDetail.Height - y
        Me.Refresh
    End If
End Sub

'因多张退费要从最后一张开始退,顺序很重要,所以不提供用户排序
'Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If mshList.MouseRow = 0 Then
'        mshList.MousePointer = 99
'    Else
'        mshList.MousePointer = 0
'    End If
'End Sub
'
'Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim lngCol As Long
'
'    lngCol = mshList.MouseCol
'
'    If Button = 1 And mshList.MousePointer = 99 Then
'        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
'        If mshList.TextMatrix(mshList.Row, col.c1类别) = "" Then Exit Sub
'        If mrsList Is Nothing Then Exit Sub
'
'        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
'        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
'
'        If mlngCurRow <> 0 Then mshList.Row = mlngCurRow
'        If mlngTopRow <> 0 Then mshList.TopRow = mlngTopRow
'    End If
'End Sub




