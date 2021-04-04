VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm按票据号批量退药 
   Caption         =   "按票据号退药"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   Icon            =   "frm按票据号批量退药.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   7725
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox Cbo批号 
      Height          =   300
      Left            =   1095
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1290
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   165
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "####"
      Top             =   1320
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   75
      ScaleHeight     =   600
      ScaleWidth      =   6765
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4140
      Width           =   6765
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   2670
         TabIndex        =   12
         ToolTipText     =   "删除当前选择行"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "退药(&O)"
         Height          =   350
         Left            =   4275
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5490
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "全清(&E)"
         Height          =   350
         Left            =   1440
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   165
         TabIndex        =   7
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.TextBox TxtNo 
      Height          =   300
      Left            =   705
      TabIndex        =   1
      Top             =   165
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill 
      Height          =   2985
      Left            =   90
      TabIndex        =   4
      Top             =   1020
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5265
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Caption         =   "处方明细"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   135
      TabIndex        =   3
      Top             =   600
      Width           =   7395
   End
   Begin VB.Label LblNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "票据号"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   540
   End
   Begin VB.Label LblNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "未输入任何处方"
      ForeColor       =   &H80000002&
      Height          =   180
      Left            =   3435
      TabIndex        =   2
      Top             =   225
      Width           =   4110
   End
End
Attribute VB_Name = "frm按票据号批量退药"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strUnit As String
Private strPrivs As String
Private mblnRefresh As Boolean
Private mlng药房ID As Long
Private mstr价格失效提示 As String
Private mint金额保留位数 As Integer                      '费用金额保留位数
Private mLngBillCount As Long

Private Enum 列名
    NO = 0
    单据
    序号
    Id
    药品ID
    分批
    批次
    科室
    姓名
    药品名称
    商品名
    规格
    批号
    效期
    产地
    付数
    数量
    已退数
    准退数
    退药数
    单位
    单价
    金额
    记录性质
    门诊标志
End Enum

Private rs序号 As New ADODB.Recordset
Private mrs退药 As New ADODB.Recordset
Private mobjPlugIn As Object             '外挂接口对象
Private Function CheckAdviceAbolish(ByVal intRow As Integer, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer) As Boolean
    '如果是退药，检查是否允许未作废医嘱退药
    Dim rstemp As ADODB.Recordset
    
    CheckAdviceAbolish = True
    On Error GoTo errHandle
    If gtype_UserSysParms.P68_门诊药嘱先作废后退药 = 0 Then Exit Function
    
    gstrSQL = "select 扣率 From 药品收发记录 Where ID=[1] "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是临嘱]", Val(Bill.TextMatrix(intRow, 列名.Id)))

    If (rstemp!扣率 Like "1*") Then       '临嘱
        gstrSQL = "Select Nvl(医嘱序号,0) 医嘱序号,Nvl(门诊标志,1) 门诊标志 From 门诊费用记录 Where ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
        If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
        Else
            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        End If
        
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是医嘱]", Val(Bill.TextMatrix(intRow, 列名.Id)))

        If Not rstemp.EOF Then
            If (rstemp!门诊标志 = 1 Or rstemp!门诊标志 = 4) And rstemp!医嘱序号 <> 0 Then
                gstrSQL = "Select Nvl(主页id, 0) As 主页id, 挂号单, decode(医嘱状态,4,1,0) 作废 From 病人医嘱记录 Where 病人来源=1  And ID=[1]"
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断该医嘱是否作废]", CLng(rstemp!医嘱序号))
                
                If Not rstemp.EOF Then
                    If rstemp!主页id > 0 And IsNull(rstemp!挂号单) Then
                        '填了主页ID，但没有挂号单的不受医嘱是否作废的限制
                    Else
                        If rstemp!作废 = 0 Then
                            MsgBox "第" & intRow & "行的药品记录对应的医嘱还未作废，不允许退药！", vbInformation, gstrSysName
                            CheckAdviceAbolish = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBillOperate() As Boolean
    Dim n As Integer
    
    For n = 1 To Bill.rows - 1
        If Bill.TextMatrix(n, 1) <> "" Then
            If CheckBillControl(4, Val(Bill.TextMatrix(n, 列名.单据)), Bill.TextMatrix(n, 列名.NO), Val(Bill.TextMatrix(n, 列名.金额))) = False Then
                Exit Function
            End If
        End If
    Next
    
    CheckBillOperate = True
End Function


Public Property Get In_权限() As String
    In_权限 = strPrivs
End Property

Public Property Let In_权限(ByVal vNewValue As String)
    strPrivs = vNewValue
End Property
Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Public Function ShowEditor(ByVal frmParent As Object, ByVal lng药房ID As Long, Optional ByVal int金额保留位数 As Integer = 2) As Boolean
    mblnRefresh = False
    mlng药房ID = lng药房ID
    mint金额保留位数 = int金额保留位数
    Me.Show 1, frmParent
    ShowEditor = mblnRefresh
End Function

Private Sub Bill_DblClick()
    
    '显示退药数文本框，缺省为当前单位格内容，允许用户修改。
    '如果输入值非法（零、空格、非法串、大于全部可退数量）则缺省为全退
    With Bill
        .Col = 列名.退药数
        If Val(.TextMatrix(Bill.Row, 列名.Id)) = 0 Then Exit Sub
        TxtInput.Tag = Val(.TextMatrix(Bill.Row, 列名.准退数))
        TxtInput.Text = Format(Val(Bill.TextMatrix(Bill.Row, 列名.退药数)), "#####0.00000;-#####0.00000; ;")
        Call ShowTxt
    End With

End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then Call Bill_DblClick
End Sub

Private Sub Bill_Scroll()
    Dim blnCancel As Boolean
    Call TxtInput_Validate(blnCancel)
    Bill.Row = Bill.TopRow
    Call Bill_EnterCell
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearAll_Click()
    Dim lngRow As Long, lngRows As Long
    '将退药数填为零
    lngRows = Bill.rows - 1
    For lngRow = 1 To lngRows
        Bill.TextMatrix(lngRow, 列名.退药数) = ""
    Next
End Sub

Private Sub cmdDelete_Click()
    Dim lngCol As Long, lngCols As Long
    If Bill.Row = Bill.rows - 1 And Bill.Row = 1 Then
        lngCols = Bill.Cols - 1
        For lngCol = 0 To lngCols
            Bill.TextMatrix(1, lngCol) = ""
        Next
    Else
        Bill.RemoveItem Bill.Row
        Call Bill_EnterCell
    End If
End Sub

Private Sub cmdOk_Click()
    Dim blnInput As Boolean
    Dim dbl退药数 As Double
    Dim StrDate As String, StrTime As String
    Dim strShow As String, strReturn As String, strSubSql As String
    Dim str批号 As String, str效期 As String, str产地 As String
    Dim lng分批 As Long, lng批次 As Long, lngRow As Long, lngRows As Long
    Dim RecRecord As New ADODB.Recordset
    Dim bln是否有退药 As Boolean
    Dim dateCurDate As Date
    Dim str药品id As String
    Dim blnIsReturn As Boolean
    Dim int门诊 As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnBeginTrans As Boolean
    Dim Int退药 As Integer
    Dim strReturnInfo As String
    Dim strReserve As String
    
    arrSql = Array()
    
    On Error GoTo ErrHand
    
    '检查是否存在数据
    lngRow = Bill.rows - 1 - IIf(Val(Bill.TextMatrix(Bill.rows - 1, 列名.Id)) = 0, 1, 0)
    If Val(Bill.TextMatrix(lngRow, 列名.药品ID)) = 0 Then Exit Sub
    Call BuildRecord
    If Not CheckCorrelation Then Exit Sub
    If Not CheckBillOperate Then Exit Sub
    
    '提示
    If MsgBox("你确定要退药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    
    mLngBillCount = 0
    LblNote.Caption = IIf(mLngBillCount = 0, "未输入任何处方", "已输入" & mLngBillCount & "张处方")
    dateCurDate = zldatabase.Currentdate()
    StrDate = Format(dateCurDate, "yyyy-MM-dd")
    StrTime = Format(dateCurDate, "hh:mm:ss")
    StrDate = StrDate & " " & StrTime
    
    Dim intUnit As Integer
    intUnit = Val(zldatabase.GetPara("药房属性", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(mlng药房ID, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(mlng药房ID, gint门诊药房)
    Else
        strUnit = GetSpecUnit(mlng药房ID, gint住院药房)
    End If
    Select Case strUnit
    Case "售价单位"
        strSubSql = "*1"
    Case "门诊单位"
        strSubSql = "*Decode(门诊包装,Null,1,0,1,门诊包装)"
    Case "住院单位"
        strSubSql = "*Decode(住院包装,Null,1,0,1,住院包装)"
    Case "药库单位"
        strSubSql = "*Decode(药库包装,Null,1,0,1,药库包装)"
    End Select
    
    lngRows = Bill.rows - 1
    For lngRow = 1 To lngRows
        If Val(Bill.TextMatrix(lngRow, 列名.退药数)) <> 0 Then
            lng分批 = Val(Bill.TextMatrix(lngRow, 列名.分批))
            lng批次 = Val(Bill.TextMatrix(lngRow, 列名.批次))
            '如果原来不分批而现在分批
            If lng批次 = 0 And lng分批 = 1 Then
                '如果批号或效期为空，则提取供用户输入
                blnInput = (Trim(Bill.TextMatrix(lngRow, 列名.批号)) = "")
                If blnInput Then
                    strShow = Bill.TextMatrix(lngRow, 列名.科室) & "||" & _
                    Bill.TextMatrix(lngRow, 列名.姓名) & "|" & Bill.TextMatrix(lngRow, 列名.药品名称) & "|" & _
                    Val(Bill.TextMatrix(lngRow, 列名.药品ID))
                    strReturn = Frm退药设置.ShowMe(Me, strShow)
                    If strReturn = "" Then Exit Sub
                    '更新批号、效期及产地
                    Bill.TextMatrix(lngRow, 列名.批号) = Split(strReturn, "|")(0)
                    Bill.TextMatrix(lngRow, 列名.效期) = Split(strReturn, "|")(1)
                    Bill.TextMatrix(lngRow, 列名.产地) = Split(strReturn, "|")(2)
                End If
            End If
        End If
    Next
    
    bln是否有退药 = False
    
    Call BuildRecordReturn
    If mrs退药.RecordCount <> 0 Then mrs退药.MoveFirst
    mrs退药.Sort = "药品ID"
    Do While Not mrs退药.EOF
        dbl退药数 = Val(Bill.TextMatrix(mrs退药!行号, 列名.退药数))
        If dbl退药数 <> 0 Then
            If CheckBill(2, Val(Bill.TextMatrix(mrs退药!行号, 列名.Id))) <> 0 Then
                Exit Sub
            End If
            
            '检查医嘱作废
            If CheckAdviceAbolish(mrs退药!行号, Val(Bill.TextMatrix(mrs退药!行号, 列名.记录性质)), Val(Bill.TextMatrix(mrs退药!行号, 列名.门诊标志))) = False Then
                Exit Sub
            End If
                            
            gstrSQL = " Select round(" & dbl退药数 & strSubSql & ",5) 数量 From 药品目录" & _
                     " Where 药品ID=[1]"
            Set RecRecord = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Bill.TextMatrix(mrs退药!行号, 列名.药品ID)))
            
            dbl退药数 = Nvl(RecRecord!数量, 0)
     
            
            str批号 = Bill.TextMatrix(mrs退药!行号, 列名.批号)
            str效期 = Bill.TextMatrix(mrs退药!行号, 列名.效期)
            str产地 = Bill.TextMatrix(mrs退药!行号, 列名.产地)
            
            blnIsReturn = False
            If CheckPrice(Val(Bill.TextMatrix(mrs退药!行号, 列名.Id)), mstr价格失效提示) = False Then
                If MsgBox("药品[" & Bill.TextMatrix(mrs退药!行号, 列名.药品名称) & "]" & mstr价格失效提示, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnIsReturn = True
                End If
            Else
                blnIsReturn = True
            End If
            
            If blnIsReturn = True Then
                '先检查或执行预调价
                Call AutoAdjustPrice_ByID(Val(mrs退药!药品ID))
            
                If Val(Bill.TextMatrix(mrs退药!行号, 列名.记录性质)) = 1 Or (Val(Bill.TextMatrix(mrs退药!行号, 列名.记录性质)) = 2 And (Val(Bill.TextMatrix(mrs退药!行号, 列名.门诊标志))) = 1 Or (Val(Bill.TextMatrix(mrs退药!行号, 列名.门诊标志))) = 4) Then
                    int门诊 = 1
                Else
                    int门诊 = 2
                End If
                
                gstrSQL = "zl_药品收发记录_部门退药("
                '收发ID
                gstrSQL = gstrSQL & Val(Bill.TextMatrix(mrs退药!行号, 列名.Id))
                '审核人
                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                '审核时间
                gstrSQL = gstrSQL & ",To_Date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')"
                '批号
                gstrSQL = gstrSQL & "," & IIf(str批号 = "", "NULL", IIf(Mid(str批号, 1, 1) = "(", "NULL", "'" & Mid(str批号, 1, 8) & "'"))
                '效期
                gstrSQL = gstrSQL & "," & IIf(str效期 = "", "NULL", "To_Date('" & Format(str效期, "yyyy-MM-dd") & "','yyyy-MM-dd')")
                '产地
                gstrSQL = gstrSQL & "," & IIf(str产地 = "", "NULL", "'" & str产地 & "'")
                '退药数
                gstrSQL = gstrSQL & "," & dbl退药数
                '退药库房
                gstrSQL = gstrSQL & ",NULL"
                '退药人
                gstrSQL = gstrSQL & ",NULL"
                '金额保留位数
                gstrSQL = gstrSQL & "," & mint金额保留位数
                '门诊
                gstrSQL = gstrSQL & "," & int门诊
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
                
                bln是否有退药 = True
                
                If InStr("," & str药品id & ",", "," & Bill.TextMatrix(mrs退药!行号, 列名.药品ID) & ",") = 0 Then
                    str药品id = IIf(str药品id = "", "", str药品id & ",") & Bill.TextMatrix(mrs退药!行号, 列名.药品ID)
                End If
                
                strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(Bill.TextMatrix(mrs退药!行号, 列名.Id)) & "," & dbl退药数
            End If
        End If
        
        mrs退药.MoveNext
    Loop
    
    '提示停用药品
    If str药品id <> "" Then
        Int退药 = 1
        Call CheckStopMedi(str药品id, Int退药)
        If Int退药 = 2 Then Exit Sub
    End If
    
    If bln是否有退药 = True Then
        '集中处理退药事务
        gcnOracle.BeginTrans
        blnBeginTrans = True
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        blnBeginTrans = False
        
        If MsgBox("你需要打印退药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_5", "ZL8_BILL_1341_5"), Me, "退药时间=" & StrDate, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), 2)
        End If
    Else
        MsgBox "本次没有退药。"
    End If
    
    '调用退药后的外挂接口
    If Not mobjPlugIn Is Nothing And bln是否有退药 Then
        On Error Resume Next
        mobjPlugIn.DrugReturnByID mlng药房ID, strReturnInfo, CDate(StrDate), strReserve
        err.Clear: On Error GoTo 0
    End If
        
    '刷新
    mblnRefresh = True
    Call SetFormat
    TxtNo.SetFocus
    Exit Sub
ErrHand:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelAll_Click()
    Dim lngRow As Long, lngRows As Long
    '将退药数填为准退数
    lngRows = Bill.rows - 1
    For lngRow = 1 To lngRows
        Bill.TextMatrix(lngRow, 列名.退药数) = Bill.TextMatrix(lngRow, 列名.准退数)
    Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call SetFormat
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With LblNote
        .Left = Me.ScaleWidth - .Width - 100
    End With
    
    With lblTitle
        .Top = TxtNo.Top + TxtNo.Height + 80
        .Width = Me.ScaleWidth - .Left - 100
    End With
    
    With picFunc
        .Left = lblTitle.Left
        .Width = lblTitle.Width
        .Top = Me.ScaleHeight - .Height
    End With
    
    With Bill
        .Left = lblTitle.Left
        .Top = lblTitle.Top + lblTitle.Height
        .Width = lblTitle.Width
        .Height = Me.ScaleHeight - picFunc.Height - .Top
    End With
    
    With cmdCancel
        .Left = picFunc.Width - .Width - 100
    End With
    With cmdOK
        .Left = cmdCancel.Left - .Width - 100
    End With
    Call Bill_Scroll
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mLngBillCount = 0
End Sub

Private Sub TxtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long, lngNewRow As Long
    Dim blnCancel As Boolean
    lngRow = Bill.Row
    lngNewRow = lngRow                  '缺省为当前行
    
    Select Case KeyCode
    Case vbKeyUp
        If Bill.Row > 1 Then lngNewRow = Bill.Row - 1
    Case vbKeyDown, vbKeyReturn
        If Bill.Row = Bill.rows - 1 Then
            Call TxtInput_Validate(blnCancel)
            cmdDelete.SetFocus
        ElseIf Bill.Row < Bill.rows - 1 Then
            lngNewRow = Bill.Row + 1
        End If
    Case Else
        Exit Sub
    End Select
    
    KeyCode = 0
    If lngRow <> lngNewRow Then
        Call TxtInput_Validate(blnCancel)
        Bill.Row = lngNewRow
        Call Bill_EnterCell
    End If
End Sub

Private Sub TxtInput_Validate(Cancel As Boolean)
    Dim blnUnValid As Boolean, dblCount As Double
    Dim rstemp As New ADODB.Recordset
    On Error Resume Next
    
    If Not TxtInput.Visible Then Exit Sub
    blnUnValid = False
    TxtInput = Trim(TxtInput)
    
    blnUnValid = (TxtInput = "")
    If Not blnUnValid Then blnUnValid = Not IsNumeric(TxtInput)
    If Not blnUnValid Then blnUnValid = Not ((Abs(TxtInput) <= Abs(TxtInput.Tag)) And ((Val(TxtInput) >= 0 And Val(TxtInput.Tag) >= 0) Or (Val(TxtInput) <= 0 And Val(TxtInput.Tag) <= 0)))
    If blnUnValid Then
        If TxtInput = "" Then
            TxtInput = 0
        Else
            TxtInput = Val(TxtInput.Tag)
        End If
    End If
    
    Bill.TextMatrix(Bill.Row, 列名.退药数) = Format(Val(TxtInput.Text), "#####0.00000;-#####0.00000; ;")
    TxtInput.Visible = False
End Sub

Private Sub TxtNo_GotFocus()
    Call zlControl.TxtSelAll(TxtNo)
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String, str单位 As String, str包装 As String
    Dim rsBill As New ADODB.Recordset
    '根据该票据号读出所有已退药
    If KeyCode <> vbKeyReturn Then Exit Sub
    strInput = Trim(UCase(TxtNo.Text))
    If strInput = "" Then Exit Sub
    
    '检查是否存在该票号
    On Error GoTo errHandle
    gstrSQL = "Select A.No " & _
             " From 票据打印内容 A,票据使用明细 B " & _
             " Where A.ID=B.打印ID And A.数据性质=1 " & _
             " And B.票种=1 And B.号码=[1]"
    Set rsBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否存在该票号]", strInput)
    
    If rsBill.RecordCount = 0 Then
        MsgBox "不存在该票据号，请重输！", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(TxtNo)
        Exit Sub
    End If
    
    Dim intUnit As Integer
    intUnit = Val(zldatabase.GetPara("药房属性", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(mlng药房ID, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(mlng药房ID, gint门诊药房)
    Else
        strUnit = GetSpecUnit(mlng药房ID, gint住院药房)
    End If
    Select Case strUnit
    Case "售价单位"
        str单位 = "X.计算单位"
        str包装 = "1"
    Case "门诊单位"
        str单位 = "D.门诊单位"
        str包装 = "D.门诊包装"
    Case "住院单位"
        str单位 = "D.住院单位"
        str包装 = "D.住院包装"
    Case "药库单位"
        str单位 = "D.药库单位"
        str包装 = "D.药库包装"
    End Select

    gstrSQL = "" & _
            " SELECT DISTINCT S.ID,S.单据,S.药品ID,S.NO,S.序号,S.扣率,P.名称 科室,C.记录性质,C.门诊标志,'' 床号,C.姓名,'['||X.编码||']'|| X.名称 As 品名,A.名称 As 商品名, " & _
            " NVL(D.药房分批,0) 分批,DECODE(X.规格,NULL,S.产地,DECODE(S.产地,NULL,X.规格,X.规格||'|'||S.产地)) 规格," & str单位 & " 单位," & str包装 & " 包装," & _
            " S.付数 付,S.实际数量 数量,S.已退数量,S.已发数量 准退数,DECODE(S.批号,NULL,'',S.批号)||DECODE(S.批次,NULL,'',0,'','('||S.批次||')') 批号," & _
            " NVL(S.批次,0) 批次,S.效期, S.零售价 单价,S.零售金额 金额,S.单量,S.频次,S.用法,S.摘要 说明,S.审核人,TO_CHAR(S.审核日期,'YYYY-MM-DD HH24:MI:SS') 发药时间,1 可操作" & _
            " FROM" & _
            "     (SELECT A.ID,A.NO,A.单据,A.序号,A.药品ID,A.费用ID,A.批次,A.批号,A.产地,A.效期,NVL(A.扣率,0) 扣率," & _
            "     NVL(A.付数,1) 付数,A.实际数量 实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
            "     A.零售价 , A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID" & _
            "     FROM 药品收发记录 A,"
    gstrSQL = gstrSQL & _
            "         (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
            "         FROM 药品收发记录 A" & _
            "         WHERE A.审核人 IS NOT NULL AND A.库房ID+0 = [1] AND A.单据 = 8" & _
            "         AND NO IN (SELECT A.NO FROM 票据打印内容 A,票据使用明细 B " & _
            "             WHERE A.ID=B.打印ID AND A.数据性质=1 AND B.票种=1 AND B.号码=[2])" & _
            "         GROUP BY A.NO,A.单据,A.药品ID,A.序号) B" & _
            "     WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 AND B.已发数量<>0 And A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)) S," & _
            "     门诊费用记录 C,部门表 P,药品规格 D,收费项目目录 X,收费项目别名 A," & _
            "     (SELECT A.NO FROM 票据打印内容 A,票据使用明细 B " & _
             "    WHERE A.ID=B.打印ID AND A.数据性质=1 AND B.票种=1 AND B.号码=[2]) B"
    gstrSQL = gstrSQL & _
            " WHERE S.药品ID=D.药品ID AND D.药品ID=X.ID AND S.对方部门ID+0=P.ID AND S.费用ID=C.ID" & _
            " AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 " & _
            " AND (S.记录状态=1 OR MOD(S.记录状态,3)=0) AND S.审核人 IS NOT NULL AND C.NO=B.NO AND S.库房ID+0=[1] " & _
            " AND S.实际数量*S.付数>S.已退数量 " & _
            " ORDER BY S.NO,S.单据"
    Set rsBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取该票据号对应的已发药记录]", mlng药房ID, strInput)
                
    If rsBill.RecordCount = 0 Then
        MsgBox "该票据号对应的处方还未发药或已全部退药或已转出到后备数据库！", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(TxtNo)
        Exit Sub
    End If
    
    Call WriteBill(rsBill)
    Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetBillID() As String
    Dim lngRow As Long, lngRows As Long
    Dim strReturn As String
    '产生已存在的处方明细ID，以备检查，如果存在相同的，则不加入
    lngRows = Bill.rows - 1
    For lngRow = 1 To lngRows
        If Val(Bill.TextMatrix(lngRow, 列名.Id)) <> 0 Then
            strReturn = strReturn & "," & Bill.TextMatrix(lngRow, 列名.Id)
        End If
    Next
    If strReturn = "" Then Exit Function
    GetBillID = strReturn & ","
End Function

Private Sub WriteBill(ByVal rsBill As ADODB.Recordset)
    Dim lngRow As Long
    Dim strID As String
    '检查起始行
    lngRow = Bill.rows - 1 + IIf(Val(Bill.TextMatrix(Bill.rows - 1, 列名.Id)) = 0, 0, 1)
    If Bill.rows - 1 < lngRow Then Bill.rows = Bill.rows + 1
    
    '产生已存在的处方明细ID，以备检查，如果存在相同的，则不加入
    strID = GetBillID
    
    '将药品明细写入已发药清单
    With rsBill
        Do While Not .EOF
            '以前没加入的记录才写在已发药清单中，供用户退药
            If InStr(1, strID, "," & !Id & ",") = 0 Then
                '因票据是当某次领用的票据全使用才一次性转出，因此，可能存在：
                '药品收发记录与费用记录已转出，而票据未转出的情况，所以此处加判断
                If Not zldatabase.NOMoved("药品收发记录", !NO, "单据=", !单据) Then
                    Bill.TextMatrix(lngRow, 列名.NO) = !NO
                    Bill.TextMatrix(lngRow, 列名.单据) = !单据
                    Bill.TextMatrix(lngRow, 列名.序号) = !序号
                    Bill.TextMatrix(lngRow, 列名.Id) = !Id
                    Bill.TextMatrix(lngRow, 列名.药品ID) = !药品ID
                    Bill.TextMatrix(lngRow, 列名.分批) = !分批
                    Bill.TextMatrix(lngRow, 列名.批次) = !批次
                    Bill.TextMatrix(lngRow, 列名.科室) = !科室
                    Bill.TextMatrix(lngRow, 列名.姓名) = !姓名
                    Bill.TextMatrix(lngRow, 列名.药品名称) = !品名
                    Bill.TextMatrix(lngRow, 列名.商品名) = IIf(IsNull(!商品名), "", !商品名)
                    Bill.TextMatrix(lngRow, 列名.规格) = IIf(IsNull(!规格), "", !规格)
                    Bill.TextMatrix(lngRow, 列名.批号) = IIf(IsNull(!批号), "", !批号)
                    Bill.TextMatrix(lngRow, 列名.效期) = IIf(IsNull(!效期), "", !效期)
                    Bill.TextMatrix(lngRow, 列名.产地) = ""
                    Bill.TextMatrix(lngRow, 列名.付数) = Format(!付, "#####0;-#####0; ;")
                    Bill.TextMatrix(lngRow, 列名.数量) = Format(!数量 / !包装, "#####0.00000;-#####0.00000; ;")
                    Bill.TextMatrix(lngRow, 列名.已退数) = Format(!已退数量 / !包装, "#####0.00000;-#####0.00000; ;")
                    Bill.TextMatrix(lngRow, 列名.准退数) = Format(!准退数 / !包装, "#####0.00000;-#####0.00000; ;")
                    Bill.TextMatrix(lngRow, 列名.退药数) = Format(!准退数 / !包装, "#####0.00000;-#####0.00000; ;")
                    Bill.TextMatrix(lngRow, 列名.单位) = IIf(IsNull(!单位), "", !单位)
                    Bill.TextMatrix(lngRow, 列名.单价) = Format(!单价 * !包装, "#####0.00000;-#####0.00000; ;")
                    Bill.TextMatrix(lngRow, 列名.单价) = !金额
                    Bill.TextMatrix(lngRow, 列名.记录性质) = !记录性质
                    Bill.TextMatrix(lngRow, 列名.门诊标志) = !门诊标志
                
                    If lngRow >= Bill.rows - 1 Then
                        lngRow = lngRow + 1
                        Bill.rows = Bill.rows + 1
                    End If
                    
                    mLngBillCount = mLngBillCount + 1
                End If
            End If
            .MoveNext
        Loop
        
        '删除最后的空白行
        If Val(Bill.TextMatrix(Bill.rows - 1, 列名.Id)) = 0 Then
            Bill.rows = Bill.rows - 1
        End If
        
        
        LblNote.Caption = IIf(mLngBillCount = 0, "未输入任何处方", "已输入" & mLngBillCount & "张处方")
    End With
End Sub

Private Sub Bill_EnterCell()
    Dim blnCancel As Boolean
    If TxtInput.Visible Then
        Call TxtInput_Validate(blnCancel)
        TxtInput.Visible = False
    End If
End Sub

Private Sub Bill_GotFocus()
    Bill_EnterCell
End Sub

Private Sub ShowTxt(Optional ByVal 对齐方式 As Integer = 1)
    '0-左对齐;1-右对齐;2-居中对齐
    On Error Resume Next
    With TxtInput
        .Alignment = 对齐方式
        .Left = Bill.Left + Bill.CellLeft
        .Top = Bill.Top + Bill.CellTop
        .Width = Bill.CellWidth - 20
        .Visible = True
        .ZOrder 0
        .SetFocus
    End With
    Call zlControl.TxtSelAll(TxtInput)
End Sub

Private Function CheckBill(ByVal IntOper As Integer, ByVal LngID As Long) As Integer
    Dim RecCheck As New ADODB.Recordset
    
    '--根据将要执行的操作，判断是否允许--
    '0-拒发;1-发药;2-退药
    '返回:
    '0-允许操作
    '1-已发药
    '2-已删除
    '3-未发药
    On Error GoTo errHandle
    gstrSQL = " Select A.审核人,Decode(Nvl(A.摘要,'小宝'),'拒发',3,B.执行状态) 执行状态 From 药品收发记录 A,门诊费用记录 B " & _
                 " Where A.费用ID=B.ID And A.ID=[1]"
        If IntOper = 2 Then
            gstrSQL = gstrSQL & " And 审核人 IS Not Null"
        Else
            gstrSQL = gstrSQL & " And 审核人 IS Null"
        End If
    Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngID)
    
    With RecCheck
        If .EOF Then CheckBill = 2: MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!审核人) Then
            If IntOper <> 2 Then CheckBill = 1: MsgBox "该处方已被其它操作员发药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
        Else
            If IntOper = 2 Then CheckBill = 3: MsgBox "该处方还未发药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
        End If
        If IntOper = 1 And !执行状态 = 3 Then CheckBill = 2: MsgBox "该处方已拒发，操作被迫中止！", vbInformation, gstrSysName: Exit Function
    End With
    
    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetFormat()
Dim intCol As Integer
    '设置表格
    With Bill
        .rows = 2
        .Cols = 25
        .Clear
        
        .TextMatrix(0, 列名.NO) = "NO"
        .TextMatrix(0, 列名.单据) = "单据"
        .TextMatrix(0, 列名.序号) = "序号"
        .TextMatrix(0, 列名.Id) = "ID"
        .TextMatrix(0, 列名.药品ID) = "药品ID"
        .TextMatrix(0, 列名.分批) = "分批"
        .TextMatrix(0, 列名.批次) = "批次"
        .TextMatrix(0, 列名.科室) = "科室"
        .TextMatrix(0, 列名.姓名) = "姓名"
        .TextMatrix(0, 列名.药品名称) = "药品名称"
        .TextMatrix(0, 列名.商品名) = "商品名"
        .TextMatrix(0, 列名.规格) = "规格"
        .TextMatrix(0, 列名.批号) = "批号"
        .TextMatrix(0, 列名.效期) = "效期"
        .TextMatrix(0, 列名.产地) = "产地"
        .TextMatrix(0, 列名.付数) = "付"
        .TextMatrix(0, 列名.数量) = "数量"
        .TextMatrix(0, 列名.已退数) = "已退数"
        .TextMatrix(0, 列名.准退数) = "准退数"
        .TextMatrix(0, 列名.退药数) = "退药数"
        .TextMatrix(0, 列名.单位) = "单位"
        .TextMatrix(0, 列名.单价) = "单价"
        .TextMatrix(0, 列名.金额) = "金额"
        .TextMatrix(0, 列名.记录性质) = "记录性质"
        .TextMatrix(0, 列名.门诊标志) = "门诊标志"
        
        .ColWidth(列名.NO) = 900
        .ColWidth(列名.单据) = 0
        .ColWidth(列名.序号) = 0
        .ColWidth(列名.Id) = 0
        .ColWidth(列名.药品ID) = 0
        .ColWidth(列名.分批) = 0
        .ColWidth(列名.批次) = 0
        .ColWidth(列名.科室) = 0
        .ColWidth(列名.姓名) = 0
        .ColWidth(列名.药品名称) = 2000
        .ColWidth(列名.规格) = 1500
        .ColWidth(列名.批号) = 1500
        .ColWidth(列名.效期) = 0
        .ColWidth(列名.产地) = 0
        .ColWidth(列名.付数) = 300
        .ColWidth(列名.数量) = 1000
        .ColWidth(列名.已退数) = 1000
        .ColWidth(列名.准退数) = 1000
        .ColWidth(列名.退药数) = 1000
        .ColWidth(列名.单位) = 600
        .ColWidth(列名.单价) = 1000
        .ColWidth(列名.金额) = 0
        .ColWidth(列名.记录性质) = 0
        .ColWidth(列名.门诊标志) = 0
    
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        .ColAlignment(列名.规格) = 1
        .ColAlignment(列名.批号) = 1
        .ColAlignment(列名.已退数) = 7
        .ColAlignment(列名.准退数) = 7
        .ColAlignment(列名.退药数) = 7
        If gint药品名称显示 = 2 Then
            If .ColWidth(列名.商品名) = 0 Then .ColWidth(列名.商品名) = 2000
        Else
            .ColWidth(列名.商品名) = 0
        End If
    End With
End Sub

Private Sub BuildRecord()
    Dim intRow As Integer, intRows As Integer
    Dim strNo As String, lng单据 As Long, str序号 As String
    
    Call InitRec
    '根据待退药清单，按单据获取明细序号
    intRows = Bill.rows - 1
    For intRow = 1 To intRows
        If Val(Bill.TextMatrix(intRow, 列名.Id)) <> 0 Then
            strNo = Bill.TextMatrix(intRow, 列名.NO)
            lng单据 = Val(Bill.TextMatrix(intRow, 列名.单据))
            If Val(Bill.TextMatrix(intRow, 列名.退药数)) <> 0 Then
                If rs序号.RecordCount <> 0 Then rs序号.MoveFirst
                rs序号.Find "单据标识='" & strNo & "|" & lng单据 & "'"
                If rs序号.EOF Then rs序号.AddNew
                rs序号!单据标识 = strNo & "|" & lng单据
                rs序号!记录性质 = Val(Bill.TextMatrix(intRow, 列名.记录性质))
                rs序号!门诊标志 = Val(Bill.TextMatrix(intRow, 列名.门诊标志))
                str序号 = Nvl(rs序号!序号)
                If InStr(1, "," & str序号 & ",", "," & Val(Bill.TextMatrix(intRow, 列名.序号)) & ",") = 0 Then
                    If str序号 = "" Then
                        str序号 = Val(Bill.TextMatrix(intRow, 列名.序号))
                    Else
                        str序号 = str序号 & "," & Val(Bill.TextMatrix(intRow, 列名.序号))
                    End If
                    rs序号!序号 = str序号
                End If
                rs序号.Update
            End If
        End If
    Next
End Sub

Private Sub BuildRecordReturn()
    '退药数据集
    Dim intRow As Integer, intRows As Integer
        
    Call InitRecReturn
    '根据待退药清单，构建退药数据集
    intRows = Bill.rows - 1
    For intRow = 1 To intRows
        If Val(Bill.TextMatrix(intRow, 列名.Id)) <> 0 Then
           If Val(Bill.TextMatrix(intRow, 列名.退药数)) <> 0 Then
                mrs退药.AddNew
                mrs退药!行号 = intRow
                mrs退药!药品ID = Val(Bill.TextMatrix(intRow, 列名.药品ID))
                mrs退药.Update
            End If
        End If
    Next
End Sub
Private Function CheckCorrelation() As Boolean
    Dim strNo As String, lng单据 As Long, str序号 As String
    '检查处方是否已结帐、检查该病人是否已出院，并对权限进行检查
    With rs序号
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !单据标识
            lng单据 = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            str序号 = Nvl(!序号)
            If Not IsReceiptBalance_Charge(1, strPrivs, lng单据, strNo, str序号, Val(!记录性质), Val(!门诊标志)) Then Exit Function
            If Not IsOutPatient(strPrivs, lng单据, strNo, Val(!记录性质), Val(!门诊标志)) Then Exit Function
            .MoveNext
        Loop
    End With
    
    CheckCorrelation = True
End Function

Private Sub InitRec()
    Set rs序号 = New ADODB.Recordset
    With rs序号
        If .State = 1 Then .Close
        .Fields.Append "单据标识", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "序号", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub InitRecReturn()
    '退药数据集，便于在批量退药时按药品ID排序
    Set mrs退药 = New ADODB.Recordset
    With mrs退药
        If .State = 1 Then .Close
        .Fields.Append "行号", adDouble, 18, adFldIsNullable
        .Fields.Append "药品id", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
