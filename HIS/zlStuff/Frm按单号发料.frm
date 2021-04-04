VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm按单号发料 
   Caption         =   "批量发料"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   Icon            =   "Frm按单号发料.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7605
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   7
      Top             =   5190
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   2580
      TabIndex        =   6
      Top             =   5190
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrintSet 
      Caption         =   "设置(&S)"
      Height          =   350
      Left            =   1350
      TabIndex        =   5
      Top             =   5190
      Visible         =   0   'False
      Width           =   1100
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   2685
      Left            =   30
      TabIndex        =   2
      Top             =   2400
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4736
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "单据明细(&D)"
      TabPicture(0)   =   "Frm按单号发料.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Msf待发明细"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "材料汇总(&T)"
      TabPicture(1)   =   "Frm按单号发料.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf待发汇总"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf待发明细 
         Height          =   2265
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3995
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483625
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf待发汇总 
         Height          =   2265
         Left            =   -74940
         TabIndex        =   9
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3995
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483625
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.TextBox TxtNo 
      Height          =   300
      Left            =   660
      TabIndex        =   0
      Top             =   180
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6420
      TabIndex        =   4
      Top             =   5190
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf待发列表 
      Height          =   1755
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   3096
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5130
      TabIndex        =   3
      Top             =   5190
      Width           =   1100
   End
   Begin VB.Label LblNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "未输入任何处方"
      ForeColor       =   &H80000002&
      Height          =   180
      Left            =   3390
      TabIndex        =   11
      Top             =   240
      Width           =   4110
   End
   Begin VB.Label LblNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "处方号"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "Frm按单号发料"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstr配料人 As String

'--外部传递参数--
Private mblnModify As Boolean
Private mstrPrivs As String
Private mint单据 As Integer                          '单据
Private mlng发料部门ID As Long                           '料房
Private mintSendAfterDosage As Integer               '允许未配料发料
Private mint允许未审核处方发料 As Integer            '允许未审核处方发料
Private mIntCheckStock As Integer                    '库存检测
Private mint校验处方 As Integer                      '校验处方
Private mintUnit  As Integer                        '单位
'--本程序使用变量--
Private mrsBill As New ADODB.Recordset              '单据记录
Private mrsTotal As New ADODB.Recordset             '汇总数据
Private mrs序号 As ADODB.Recordset
Private mrs处方来源部门 As ADODB.Recordset            '记录所有待发料处方的来源部门
Private mrs待发汇总明细 As ADODB.Recordset            '记录待发汇总的记录，实际是按单据号的明细记录


Private mblnStartUp As Boolean
Private mlngListRow As Long                          '待发列表
Private mlngDetailRow As Long                        '待发明细
Private mlngTotalRow As Long                         '待发汇总
Private mstrBillNo As String                         '汇总单据号
Private mstrID As String                             '汇总ID
Private mlngBillCount As Long
Private mstr单据号 As String
Private mint单据类型 As String
Private mstr单据IN  As String
Private mbln按票据号发料 As Boolean
Private Const mlngModule = 1723
Private mobjPlugIn As Object             '外挂接口对象
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Public Property Get 按票据号发料() As Boolean
   按票据号发料 = mbln按票据号发料
End Property

Public Property Let 按票据号发料(ByVal vNewValue As Boolean)
    mbln按票据号发料 = vNewValue
End Property

Public Property Get In_权限() As String
    In_权限 = mstrPrivs
End Property

Public Property Let In_权限(ByVal vNewValue As String)
    mstrPrivs = vNewValue
End Property
Public Property Get In_单据IN() As String
    In_单据IN = mstr单据IN
End Property

Public Property Let In_单据IN(ByVal vNewValue As String)
    mstr单据IN = vNewValue
End Property


Public Property Get In_单据() As Integer
    In_单据 = mint单据
End Property

Public Property Let In_单据(ByVal vNewValue As Integer)
    mint单据 = vNewValue
End Property

Public Property Get In_校验处方() As Integer
    In_校验处方 = mint校验处方
End Property

Public Property Let In_校验处方(ByVal vNewValue As Integer)
    mint校验处方 = vNewValue
End Property

Public Property Get In_库存检查() As Integer
    In_库存检查 = mIntCheckStock
End Property

Public Property Let In_库存检查(ByVal vNewValue As Integer)
    mIntCheckStock = vNewValue
End Property

Public Property Get In_发料部门id() As Long
    In_发料部门id = mlng发料部门ID
End Property

Public Property Let In_发料部门id(ByVal vNewValue As Long)
    mlng发料部门ID = vNewValue
End Property

Public Property Get In_允许未配料发料() As Integer
    In_允许未配料发料 = mintSendAfterDosage
End Property

Public Property Let In_允许未配料发料(ByVal vNewValue As Integer)
    mintSendAfterDosage = vNewValue
End Property

Public Property Get IN_允许未审核发料() As Integer
    IN_允许未审核发料 = mint允许未审核处方发料
End Property

Public Property Let IN_允许未审核发料(ByVal vNewValue As Integer)
    mint允许未审核处方发料 = vNewValue
End Property

Private Sub SetFormat(Optional ByVal IntStyle As Integer = 1)
    Dim intCol As Integer
    '设置各列表控件的格式

    Select Case IntStyle
    Case 1
        With Msf待发列表
            .Rows = 2
            .Cols = 10
    
            .TextMatrix(0, 0) = "类型"
            .TextMatrix(0, 1) = "NO"
            .TextMatrix(0, 2) = "科室"
            .TextMatrix(0, 3) = "姓名"
            .TextMatrix(0, 4) = "住院号"
            .TextMatrix(0, 5) = "床号"
            .TextMatrix(0, 6) = "收费员"
            .TextMatrix(0, 7) = "开单医生"
            .TextMatrix(0, 8) = "开单日期"
            .TextMatrix(0, 9) = "门诊标志"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If mblnStartUp = False Then
                .ColWidth(0) = 500
                .ColWidth(1) = 1000
                .ColWidth(2) = 1200
                .ColWidth(3) = 1000
                .ColWidth(4) = 1000
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 1000
                .ColWidth(8) = 1200
                .ColWidth(9) = 0
                
                .Row = 1
                Call RestoreFlexState(Msf待发列表, Me.Name)
                If glngSys \ 100 <> 1 Then
                    .ColWidth(2) = 0
                    .ColWidth(4) = 0
                    .ColWidth(5) = 0
                End If
                .ColWidth(7) = IIf(mint校验处方 = 1, 0, 1000)
            End If
        End With
    Case 2
        With Msf待发明细
            .Rows = 2
            .Cols = 6
    
            .TextMatrix(0, 0) = "材料名称"
            .TextMatrix(0, 1) = "规格"
            .TextMatrix(0, 2) = "单位"
            .TextMatrix(0, 3) = "单价"
            .TextMatrix(0, 4) = "数量"
            .TextMatrix(0, 5) = "金额"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
                If intCol < 2 Then .ColAlignment(intCol) = 1
                If intCol > 2 Then .ColAlignment(intCol) = 7
            Next
    
            If mblnStartUp = False Then
                .ColWidth(0) = 2000
                .ColWidth(1) = 1500
                .ColWidth(2) = 500
                .ColWidth(3) = 800
                .ColWidth(4) = 800
                .ColWidth(5) = 1000
                
                .Row = 1
                Call RestoreFlexState(Msf待发明细, Me.Name)
            End If
        End With
    Case 3
        With Msf待发汇总
            .Rows = 2
            .Cols = 9
    
            .TextMatrix(0, 0) = "序号"
            .TextMatrix(0, 1) = "材料名称"
            .TextMatrix(0, 2) = "规格"
            .TextMatrix(0, 3) = "单位"
            .TextMatrix(0, 4) = "单价"
            .TextMatrix(0, 5) = "数量"
            .TextMatrix(0, 6) = "金额"
            .TextMatrix(0, 7) = "材料ID"
            .TextMatrix(0, 8) = "批次"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
                If intCol < 3 Then .ColAlignment(intCol) = 1
                If intCol > 3 Then .ColAlignment(intCol) = 7
                
            Next
            
            If mblnStartUp = False Then
                .ColWidth(0) = 500
                .ColWidth(1) = 2000
                .ColWidth(2) = 1500
                .ColWidth(3) = 500
                .ColWidth(4) = 800
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 0
                .ColWidth(8) = 0
                .Row = 1
                Call RestoreFlexState(Msf待发汇总, Me.Name)
            End If
        End With
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    If CheckStock = False Then Exit Sub
    If Not CheckCorrelation Then Exit Sub
    If SendBill = False Then Exit Sub
    
    mlngBillCount = 0
    lblNote.Caption = IIf(mlngBillCount = 0, "未输入任何处方", "已输入" & mlngBillCount & "张处方")
    
    '初始化
    mstrID = ""
    mstrBillNo = ""
    txtNO = ""
    
    With Msf待发汇总
        .Clear
        .Rows = 2
        .RowData(1) = 0
    End With
    With Msf待发列表
        .Clear
        .Rows = 2
        .RowData(1) = 0
    End With
    With Msf待发明细
        .Clear
        .Rows = 2
        .RowData(1) = 0
    End With
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
    cmdOK.Enabled = False
    txtNO.SetFocus
End Sub

Private Sub CmdPrint_Click()
    Dim HisPrint As New zlPrint1Grd
    Dim HisRow As New zlTabAppRow
    Dim ArrayNo, IntArray As Integer
    Dim LngSelectRow As Long, intCol As Integer
    
    On Error Resume Next
    '取消表格的选择状态
    With Msf待发汇总
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If mlngTotalRow > 0 And mlngTotalRow < .Rows Then
            .Row = mlngTotalRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
    End With
    
    HisPrint.Title = "材料汇总"
    Set HisRow = New zlTabAppRow
    HisRow.Add "日期:" & Format(Sys.Currentdate, "yyyy年MM月dd日")
    HisPrint.UnderAppRows.Add HisRow
    
    ArrayNo = Split(mstrBillNo, ";")
    
    Set HisRow = New zlTabAppRow
    HisRow.Add "单据号:"
    HisPrint.BelowAppRows.Add HisRow
    For IntArray = 0 To UBound(ArrayNo)
        Set HisRow = New zlTabAppRow
        HisRow.Add Space(10) & ArrayNo(IntArray)
        HisPrint.BelowAppRows.Add HisRow
    Next
    
    Set HisPrint.Body = Msf待发汇总
    Select Case zlPrintAsk(HisPrint)
    Case 1
        zlPrintOrView1Grd HisPrint, 1
    Case 2
        zlPrintOrView1Grd HisPrint, 2
    Case 3
        zlPrintOrView1Grd HisPrint, 3
    End Select
    
    '恢复表格的选择状态
    With Msf待发汇总
        
        mlngTotalRow = LngSelectRow
        .Row = mlngTotalRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub cmdPrintSet_Click()
    zlPrintSet
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strReg As String
    
    mblnStartUp = False
    mlngBillCount = 0
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
   ' mbln按票据号发料 = False
    If mbln按票据号发料 = True Then lblNO.Caption = "票据号": Me.Caption = "按票据号发料"
    mstrID = ""
    mstrBillNo = ""
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
    mblnStartUp = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < 7730 Then Me.Width = 7730
    If Me.Height < 6000 Then Me.Height = 6000
    
    With lblNote
        .Left = Me.ScaleWidth - .Width - 100
    End With
    
    With cmdHelp
        .Top = Me.ScaleHeight - .Height - 100
    End With
    With cmdPrintSet
        .Top = cmdHelp.Top
        .Left = cmdHelp.Left + cmdHelp.Width + 100
    End With
    With cmdPrint
        .Top = cmdHelp.Top
        .Left = cmdPrintSet.Left + cmdPrintSet.Width + 100
    End With
    
    With CmdCancel
        .Top = cmdHelp.Top
        .Left = Me.ScaleWidth - .Width - 100
    End With
    With cmdOK
        .Top = cmdHelp.Top
        .Left = CmdCancel.Left - .Width - 100
    End With
    
    With Msf待发列表
        .Height = (cmdOK.Top - 200 - .Top) / 2
        .Width = Me.ScaleWidth - .Left - 50
    End With
    
    With TabShow
        .Top = Msf待发列表.Top + Msf待发列表.Height + 100
        .Height = cmdOK.Top - 100 - .Top
        .Width = Msf待发列表.Width
    End With
    With Msf待发汇总
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 50
    End With
    With Msf待发明细
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 50
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(Msf待发汇总, Me.Name)
    Call SaveFlexState(Msf待发列表, Me.Name)
    Call SaveFlexState(Msf待发明细, Me.Name)
End Sub

Private Sub Msf待发汇总_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf待发汇总
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If mlngTotalRow > 0 And mlngTotalRow < .Rows Then
            .Row = mlngTotalRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        mlngTotalRow = LngSelectRow
        .Row = mlngTotalRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Msf待发汇总_GotFocus()
    With Msf待发汇总
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf待发汇总_LostFocus()
    With Msf待发汇总
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf待发列表_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf待发列表
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If mlngListRow > 0 And mlngListRow < .Rows Then
            .Row = mlngListRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        mlngListRow = LngSelectRow
        .Row = mlngListRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
        
        If Trim(.TextMatrix(.Row, 1)) = "" Then
            With Msf待发明细
                .Clear
                .Rows = 2
                Call SetFormat(2)
            End With
            Exit Sub
        End If
        
        '显示单据明细
        Call ReadBillData(.RowData(.Row), .TextMatrix(.Row, 1), Val(.TextMatrix(.Row, 9)))
    End With
End Sub

Private Sub Msf待发列表_GotFocus()
    With Msf待发列表
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf待发列表_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng单据 As Long, strNo As String
    
    If KeyCode = vbKeyDelete Then
        If Msf待发列表.TextMatrix(Msf待发列表.Row, 1) = "" Then Exit Sub
        
        With mrs待发汇总明细
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    .Find "单据号='" & Msf待发列表.TextMatrix(Msf待发列表.Row, 1) & "'"
                    If Not .EOF Then .Delete
                    If Not .EOF Then .MoveNext
                Loop
            End If
        End With
        With mrs处方来源部门
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "来源部门='" & Msf待发列表.TextMatrix(Msf待发列表.Row, 2) & "'"
                If Not .EOF Then .Delete
            End If
        End With
        With Msf待发列表
            lng单据 = Val(.TextMatrix(.Row, 0))
            strNo = .TextMatrix(.Row, 1)
            If .Rows - 1 = 1 Then
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
                .TextMatrix(1, 6) = ""
                .RowData(1) = 0
            Else
                If Trim(.TextMatrix(.Row, 1)) <> "" Then .RemoveItem .Row: mlngBillCount = mlngBillCount - 1
            End If
            
            cmdOK.Enabled = (.RowData(IIf(.Rows - 1 = 1, 1, .Rows - 2)) <> 0)
            lblNote.Caption = IIf(mlngBillCount = 0, "未输入任何处方", "已输入" & mlngBillCount & "张处方")
        
            '删除该单据
            With mrs序号
                If .RecordCount <> 0 Then .MoveFirst
                .Find "单据标识='" & strNo & "|" & lng单据 & "'"
                If Not .EOF Then .Delete
            End With
            
        End With
        
        Msf待发列表_EnterCell
        mblnModify = True
        Call WriteTotalDataToBill
    End If
End Sub

Private Sub Msf待发列表_LostFocus()
    With Msf待发列表
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf待发明细_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf待发明细
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If mlngDetailRow > 0 And mlngDetailRow < .Rows Then
            .Row = mlngDetailRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        mlngDetailRow = LngSelectRow
        .Row = mlngDetailRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Msf待发明细_GotFocus()
    With Msf待发明细
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf待发明细_LostFocus()
    With Msf待发明细
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    Select Case TabShow.Tab
    Case 0
        Msf待发明细.ZOrder
        Msf待发明细_EnterCell
    Case 1
        WriteTotalDataToBill
        Msf待发汇总.ZOrder
        Msf待发汇总_EnterCell
    End Select
End Sub

Private Sub TxtNo_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub

Private Function Send处方号() As Boolean
    '--------------------------------------------------------------------------------------------------------------------------------------------
    '--功能:按处方号发料
    '--参数:
    '--返回:发料成功,返回true,否则返回False
    '--------------------------------------------------------------------------------------------------------------------------------------------

    Dim intYear As Integer, strYear As String
    Dim intRow As Integer
    Dim strNo As String, IntBill As Integer, ArrTmp, strTmp As String
    Dim strSql As String
    Dim int门诊 As Integer
    
    '--如果不满八位,则按规则产生--
    Me.txtNO = UCase(LTrim(Me.txtNO))
    Me.txtNO.Text = zlCommFun.GetFullNO(Me.txtNO.Text, 13)
    On Error GoTo ErrHandle
    gstrSQL = "" & _
             " Select /*+ Rule*/ Distinct Decode(C.单据,24,'收费',25,'记帐',26,'记帐表') 类型,C.No,C.单据,A.已收费," & _
             "      Decode(A.配药人,Null,'','部门发料','',A.配药人) 配料人,P.名称 科室,decode(c.单据,26,'',B.姓名) 姓名," & _
             "      Decode(c.单据,26,'',B.标识号)  住院号,decode(c.单据,26,'','') 床号,B.开单人 开单医生,B.操作员姓名 填制人," & _
             "      To_Char(C.填制日期,'yyyy-MM-dd') 填制日期,0 门诊 " & _
             " From 未发药品记录 A,门诊费用记录 B,药品收发记录 C,部门表 P,部门表 S " & _
             "     ,Table(cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
             " Where C.费用ID=B.ID And B.开单部门ID+0=P.ID(+) And Nvl(C.库房ID,0)+0=S.ID(+) " & _
             "     And Nvl(A.库房ID,0)=Nvl(C.库房ID,0) And Mod(C.记录状态,3)=1 And A.No=C.No " & _
             "     And (C.库房ID+0=[2] OR C.库房ID IS NULL)" & _
             "     And C.单据=D.Column_Value And C.审核人 Is Null " & _
             "     And C.单据=A.单据 And C.No=[1] and nvl(C.发药方式,-999)<>-1 And Nvl(B.费用状态,0)<>1 "
     
    If mstr单据IN = "24" Then
    ElseIf mstr单据IN = "26" Then
        gstrSQL = Replace(gstrSQL, "0 门诊", "1 门诊")
        gstrSQL = Replace(gstrSQL, "B.姓名", "nvl(R.姓名,B.姓名)")
        gstrSQL = Replace(gstrSQL, "decode(c.单据,26,'','') 床号", "decode(c.单据,26,'',B.床号) 床号")
        gstrSQL = Replace(gstrSQL, "门诊费用记录 B", "住院费用记录 B,病案主页 R")
        gstrSQL = Replace(gstrSQL, "And Nvl(B.费用状态,0)<>1", "And B.病人id=R.病人id And B.主页id=R.主页id")
    ElseIf InStr(1, mstr单据IN, "25") > 0 Or InStr(1, mstr单据IN, "26") > 0 Then
        strSql = Replace(gstrSQL, "0 门诊", "1 门诊")
        strSql = Replace(strSql, "B.姓名", "nvl(R.姓名,B.姓名)")
        strSql = Replace(strSql, "decode(c.单据,26,'','') 床号", "decode(c.单据,26,'',B.床号) 床号")
        strSql = Replace(strSql, "门诊费用记录 B", "住院费用记录 B,病案主页 R")
        strSql = Replace(strSql, "And Nvl(B.费用状态,0)<>1", "And B.病人id=R.病人id And B.主页id=R.主页id")
        gstrSQL = gstrSQL & " Union All " & strSql
    End If
    
'    err = 0: On Error Resume Next
    Set mrsBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtNO, mlng发料部门ID, mstr单据IN)
    If err <> 0 Then
        MsgBox "读取处方时，发生不可预知的错误！", vbInformation, gstrSysName
        GoTo ExitSub  '读取所有未发料记录
    End If
'    If ReadData(gstrSQL) = False Then GoTo ExitSub  '读取所有未发料记录

    If mrsBill.EOF Then
        MsgBox "未找到指定处方，请重新输入！", vbInformation, gstrSysName
        GoTo ExitSub
    End If
    
    If mrsBill.RecordCount > 1 Then
        strTmp = Frm单据选择.ShowMe(Me, mrsBill)
        If strTmp = "" Then GoTo ExitSub
        
        ArrTmp = Split(strTmp, ";")
        strNo = ArrTmp(0)
        IntBill = ArrTmp(1)
        
        mrsBill.MoveFirst
        mrsBill.Find "单据=" & IntBill & " And No=" & strNo
        int门诊 = mrsBill!门诊
    Else
        strNo = mrsBill!NO
        IntBill = mrsBill!单据
        int门诊 = mrsBill!门诊
    End If
    Me.txtNO.Tag = IntBill
    
    '如果已存在该单据，则退出
    If SetLocateBill(txtNO.Text, IntBill, False) Then
        MsgBox "该处方已经输入，请重输！", vbInformation, gstrSysName
        GoTo ExitSub
    End If
    
    '检测合法性
    If CheckBill(IntBill, strNo) <> 0 Then GoTo ExitSub
    '如果当前输入处方的科室与已录入的处方的科室不同，则给予提示
    If CheckSource(IntBill, strNo) = False Then Exit Function
    If WriteSendListData(IntBill, strNo, int门诊) = False Then GoTo ExitSub
    
    mlngBillCount = mlngBillCount + 1
    lblNote.Caption = IIf(mlngBillCount = 0, "未输入任何处方", "已输入" & mlngBillCount & "张处方")
    
    '定位到刚才输入的处方单
    Call SetLocateBill(txtNO.Text, Val(txtNO.Tag))
    
    With Msf待发列表
        cmdOK.Enabled = (.RowData(IIf(.Rows - 1 = 1, 1, .Rows - 2)) <> 0)
    End With
    
    mblnModify = True
    Call RefreshData
    With txtNO
        .SelStart = 0
        .SelLength = Len(txtNO)
    End With
    Send处方号 = True
    Exit Function
ExitSub:
    With txtNO
        .SelStart = 0
        .SelLength = Len(txtNO)
        .SetFocus
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtNO) = "" Then Exit Sub
    If 按票据号发料 = True Then
        If Send发票号 = False Then Exit Sub
    Else
        If Send处方号() = False Then Exit Sub
    End If
    
End Sub
Private Function Send发票号() As Boolean
    Dim blnAdd As Boolean
    
    Dim strNo As String, IntBill As Integer
    Dim rs票据 As New ADODB.Recordset
    txtNO.Text = Trim(UCase(txtNO.Text))
    
    On Error GoTo ErrHandle
    '根据输入的票据号提取处方
    gstrSQL = "Select Distinct A.No " & _
             " From 票据打印内容 A,票据使用明细 B " & _
             " Where A.ID=B.打印ID And A.数据性质=1 " & _
             " And B.票种=1 And B.号码=[1]"
    Set rs票据 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[根据输入的票据号提取处方]", txtNO.Text)
    
    If rs票据.RecordCount = 0 Then
        MsgBox "没有找到任何处方！", vbInformation, gstrSysName
        GoTo ExitSub
        Exit Function
    End If
    
    With rs票据
        Do While Not .EOF
            gstrSQL = " Select Distinct Decode(C.单据,24,'收费','记帐') 类型,C.No,C.单据,A.已收费,Decode(A.配药人,Null,'','部门发药','',A.配药人) 配药人,P.名称 科室,B.姓名,B.标识号 住院号,'' 床号,B.开单人 开单医生,B.操作员姓名 填制人,To_Char(C.填制日期,'yyyy-MM-dd') 填制日期,0 门诊 " & _
                      " From 未发药品记录 A,门诊费用记录 B,药品收发记录 C,部门表 P,部门表 S " & _
                      " Where C.费用ID=B.ID And B.开单部门ID+0=P.ID(+) And Nvl(C.库房ID,0)+0=S.ID(+) and Nvl(A.库房ID,0)=Nvl(C.库房ID,0) And Mod(C.记录状态,3)=1 And A.No=C.No " & _
                      "     And (C.库房ID+0=[2] OR C.库房ID IS NULL)" & _
                      "     And C.单据 =24 And C.审核人 Is Null And C.单据=A.单据 And C.No=[1] and nvl(C.发药方式,-999)<>-1 "
                  
            Set mrsBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(Nvl(!NO)), mlng发料部门ID)
            blnAdd = (mrsBill.RecordCount <> 0)
            
            If blnAdd Then     '找到指定处方
                strNo = mrsBill!NO
                IntBill = mrsBill!单据
                txtNO.Tag = IntBill
                
                '如果已存在该单据，则退出
                blnAdd = Not SetLocateBill(strNo, IntBill, False)
                '检测合法性
                If blnAdd Then blnAdd = Not (CheckBill(IntBill, strNo) <> 0)
                If blnAdd Then blnAdd = WriteSendListData(IntBill, strNo, 0)
                If blnAdd Then
                    mlngBillCount = mlngBillCount + 1
                    lblNote.Caption = IIf(mlngBillCount = 0, "未输入任何处方", "已输入" & mlngBillCount & "张处方")
                End If
            End If
            .MoveNext
        Loop
    End With
    
    '定位到刚才输入的处方单
    Call SetLocateBill(strNo, True)
    
    With Msf待发列表
        cmdOK.Enabled = (.RowData(IIf(.Rows - 1 = 1, 1, .Rows - 2)) <> 0)
    End With
    mblnModify = True
    Call RefreshData
    Send发票号 = True
    Exit Function
ExitSub:
    With txtNO
        .SelStart = 0
        .SelLength = Len(txtNO)
        .SetFocus
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckSource(ByVal int单据 As Integer, ByVal strNo As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim bln重复部门 As Boolean
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select B.编码 as 编码,B.名称 as 来源部门 " & _
        "   From 药品收发记录 A,部门表 B " & _
        "   Where A.对方部门id=B.id and No=[1] And 单据=[2]" & _
        "           And Mod(记录状态,3)=1 And 审核人 Is Null And (库房ID+0=[3] Or 库房ID Is NULL) And Rownum<2"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "检查", strNo, int单据, mlng发料部门ID)
    
    
    If rs.RecordCount = 0 Then
        CheckSource = False
        Exit Function
    End If
    
    With mrs处方来源部门
        If .RecordCount = 0 Then
            .AddNew
            !编码 = rs!编码
            !来源部门 = rs!来源部门
            CheckSource = True
        Else
            .MoveFirst
            For n = 1 To .RecordCount
                If !编码 = rs!编码 Then
                    bln重复部门 = True
                    Exit For
                End If
                .MoveNext
            Next
            If Not bln重复部门 Then
                If MsgBox("当前处方的开单科室是[" & rs!编码 & "]" & rs!来源部门 & "，你确定要加入该处方吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                Else
                    .AddNew
                    !编码 = rs!编码
                    !来源部门 = rs!来源部门
                    CheckSource = True
                End If
            Else
                CheckSource = True
            End If
        End If
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadData(ByVal StrQuery As String) As Boolean
    '--读取数据--

'    On Error Resume Next
'    err = 0
    On Error GoTo ErrHandle
    ReadData = False

    gstrSQL = StrQuery
    Call zlDatabase.OpenRecordset(mrsBill, gstrSQL, Me.Caption)
    If err <> 0 Then
        MsgBox "读取处方时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    ReadData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadBillData(ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int门诊 As Integer) As Boolean
    Dim IntStyle As Integer
    Dim str序号 As String
    Dim str明细单位串 As String, str汇总单位串 As String
    '--读取单据内容--
    'BillStyle-单据类型;BIllNO-单据号
    '单位显示根据服务对象来（门诊：门诊单位；住院或住院门诊：住院单位；其它；售价单位）
'    On Error Resume Next
'    err = 0
    On Error GoTo ErrHandle
    ReadBillData = False
    
    
    Select Case mintUnit
    Case 0
        str明细单位串 = "C.计算单位 单位,B.零售价 单价,B.实际数量*Nvl(B.付数,1) 数量"
        str汇总单位串 = "C.计算单位 单位,B.零售价 单价,Sum(B.实际数量*Nvl(B.付数,1)) 数量"
    Case Else
        str明细单位串 = "D.包装单位 单位,B.零售价*nvl(D.换算系数,1) 单价,B.实际数量/nvl(D.换算系数,1)*Nvl(B.付数,1) 数量"
        str汇总单位串 = "D.包装单位 单位,B.零售价*nvl(D.换算系数,1) 单价,Sum(B.实际数量/nvl(D.换算系数,1)*Nvl(B.付数,1)) 数量"
    End Select
    
    str明细单位串 = str明细单位串 & ",B.零售金额 金额 "
    str汇总单位串 = str汇总单位串 & ",Sum(B.零售金额) 金额 "

    gstrSQL = "" & _
        "   SELECT DISTINCT F.序号,F.病人ID,'['||C.编码||']'||C.名称  As 品名,DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地)) 规格," & _
                str明细单位串 & _
        " FROM 药品收发记录 B,材料特性 D,收费项目目录 C,门诊费用记录 F" & _
        " WHERE B.药品ID=D.材料ID AND D.材料ID=C.ID And B.费用ID=F.ID" & _
        "       AND MOD(B.记录状态,3)=1 AND B.NO=[1] AND B.单据=[2]" & _
        "       AND (B.库房ID+0=[3] OR B.库房ID IS NULL)"
    gstrSQL = gstrSQL & " And b.审核人 Is Null"
    
    If int门诊 = 1 Then
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    End If
    
    gstrSQL = gstrSQL & " Order by 序号"

    Set mrsBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, BillNo, BillStyle, mlng发料部门ID)

    With mrsBill
        str序号 = ""
        Do While Not .EOF
            str序号 = str序号 & "," & !序号
            .MoveNext
        Loop
        If str序号 <> "" Then str序号 = Mid(str序号, 2)
        .MoveFirst
    End With
    
    
    '将单据信息与明细序号写入内部映射记录集中
    With mrs序号
        If .RecordCount <> 0 Then .MoveFirst
        .Find "单据标识='" & BillNo & "|" & BillStyle & "'"
        If str序号 <> "" Then
            If .EOF Then
                .AddNew
                !单据标识 = BillNo & "|" & BillStyle
                !序号 = str序号
                .Update
            End If
        End If
    End With
    
    If WriteDataToBill() = False Then Exit Function

    If err <> 0 Then
        MsgBox "读取处方时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    ReadBillData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBill(ByVal intBillStyle As Integer, ByVal strNo As String) As Integer
    Dim rsCheck As New ADODB.Recordset

    '--根据将要执行的操作，判断是否允许--
    '返回:
    '0-允许操作
    '1-未配料
    '2-已配料
    '3-已发料
    '4-已删除
    '5-未发料
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select A.配药人 配料人,A.审核人,nvl(B.已收费,0) 已收费 " & _
        "   From 药品收发记录 A,未发药品记录 B " & _
        "  Where A.No=B.No And A.单据=B.单据 And A.No=[1] And A.单据=[2]" & _
        "           And mod(A.记录状态,3)=1 And Rownum=1 And (A.库房ID+0=[3] Or A.库房ID Is NULL)"
    gstrSQL = gstrSQL & " And A.审核人 IS Null"
    
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, intBillStyle, mlng发料部门ID)
    
    With rsCheck
        If .EOF Then CheckBill = 4: MsgBox "未找到处方[" & strNo & "],可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!审核人) Then
            CheckBill = 3: MsgBox "该处方[" & strNo & "]已被其它操作员发料，发料操作中止！", vbInformation, gstrSysName: Exit Function
        End If
'        If frm卫材发放管理.mint允许未审核处方发料 = 0 Then
'            If !已收费 = 0 Then
'                CheckBill = 3: MsgBox "该处方[" & strNo & "]还未收费，发料操作中止！", vbInformation, gstrSysName: Exit Function
'            End If
'        End If
    End With

    CheckBill = 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteSendListData(ByVal int单据 As Integer, ByVal strNo As String, ByVal int门诊 As Integer) As Boolean
    Dim rsCheck As New ADODB.Recordset
'    On Error Resume Next
'    err = 0
    On Error GoTo ErrHandle
    WriteSendListData = False
    
    If mintSendAfterDosage = 0 Then
        If IsNull(mrsBill!配料人) Then
            MsgBox "该处方还未配料，不能执行发料操作！", vbInformation, gstrSysName
            Exit Function
        End If
        If Trim(mrsBill!配料人) = "" Then
            MsgBox "该处方还未配料，不能执行发料操作！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mint允许未审核处方发料 = 0 Then
        If mrsBill!已收费 = 0 Then
            MsgBox "该处方还未收费或记帐，不能执行发料操作！", vbInformation, gstrSysName
            Exit Function
        End If
        
        gstrSQL = "Select 操作员姓名 " & _
            "   From 门诊费用记录 " & _
            "   Where ID =( Select 费用ID From 药品收发记录 Where 审核人 Is Null And Mod(记录状态,3)=1  And No=[1] And 单据=[2] And Rownum=1)"
        
        If int门诊 = 1 Then
            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        End If
        
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, int单据)
        
        With rsCheck
            If IsNull(!操作员姓名) Then
                MsgBox "该处方还未审核，不能执行发料操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End With
    End If
    
    With Msf待发列表
        .Redraw = False
        .TextMatrix(.Rows - 1, 0) = mrsBill!类型
        .TextMatrix(.Rows - 1, 1) = mrsBill!NO
        .TextMatrix(.Rows - 1, 2) = IIf(IsNull(mrsBill!科室), "", mrsBill!科室)
        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(mrsBill!姓名), "", mrsBill!姓名)
        .TextMatrix(.Rows - 1, 4) = IIf(IsNull(mrsBill!住院号), "", mrsBill!住院号)
        .TextMatrix(.Rows - 1, 5) = IIf(IsNull(mrsBill!床号), "", mrsBill!床号)
        .TextMatrix(.Rows - 1, 6) = IIf(IsNull(mrsBill!填制人), "", mrsBill!填制人)
        .TextMatrix(.Rows - 1, 7) = IIf(IsNull(mrsBill!开单医生), "", mrsBill!开单医生)
        .TextMatrix(.Rows - 1, 8) = IIf(IsNull(mrsBill!填制日期), "", mrsBill!填制日期)
        .TextMatrix(.Rows - 1, 9) = mrsBill!门诊
        .RowData(.Rows - 1) = mrsBill!单据
        mstr单据号 = mrsBill!NO
        mint单据类型 = mrsBill!单据
    End With
    
    If err <> 0 Then
        MsgBox "写单据列表时，发生不可预知的错误！", vbInformation, gstrSysName
        With Msf待发列表
            If .Rows - 1 >= 2 Then
                .Rows = .Rows - 1
            Else
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = ""
                .TextMatrix(.Rows - 1, 2) = ""
                .TextMatrix(.Rows - 1, 3) = ""
                .TextMatrix(.Rows - 1, 4) = ""
                .TextMatrix(.Rows - 1, 5) = ""
                .TextMatrix(.Rows - 1, 6) = ""
                .TextMatrix(.Rows - 1, 7) = ""
                .TextMatrix(.Rows - 1, 8) = ""
                .TextMatrix(.Rows - 1, 9) = ""
                .RowData(.Rows - 1) = 0
            End If
            .Redraw = True
        End With
        Exit Function
    End If
    
    With Msf待发列表
        .Rows = .Rows + 1
        .RowData(.Rows - 1) = 0
        .Redraw = True
    End With
    
    WriteSendListData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function RefreshData() As Boolean
    Dim intRow As Integer, intRows As Integer
    Dim arrID
    Dim strNoThis As String, intBillThis As Integer
    Dim str明细单位串 As String, str汇总单位串 As String
    
    On Error GoTo ErrHandle
    If mblnModify = False Then Exit Function
    RefreshData = False
    
    '清空汇总表格
    With Msf待发汇总
        .Clear
        .Rows = 2
        SetFormat (3)
    End With
    
    gstrSQL = ""
    If mbln按票据号发料 Then
            With Msf待发列表
                    '.TextMatrix(0, 0) = "类型"
                    '.TextMatrix(0, 1) = "NO"
                    '.TextMatrix(0, 2) = "科室"
                    '.TextMatrix(0, 3) = "姓名"
                    '.TextMatrix(0, 4) = "住院号"
                    '.TextMatrix(0, 5) = "床号"
                    '.TextMatrix(0, 6) = "收费员"
                    '.TextMatrix(0, 7) = "开单医生"
                    '.TextMatrix(0, 8) = "开单日期"
                For intRow = 1 To .Rows - 1
                            
                    If Trim(.TextMatrix(intRow, 1)) <> "" Then
                            '构建SQL语句
                            gstrSQL = gstrSQL & " UNION ALL  SELECT " & .RowData(intRow) & " as 单据,'" & Trim(.TextMatrix(intRow, 1)) & "' as NO From DUAL" & vbCrLf
                    End If
                Next
            End With
            
        If gstrSQL = "" Then Exit Function
        gstrSQL = Mid(gstrSQL, Len(" UNION ALL "))
        gstrSQL = "" & _
            "   Select NO,药品ID,批次,零售价,实际数量,付数,零售金额 " & _
            "   From 药品收发记录 " & _
            "   Where (单据,No) in (" & gstrSQL & ") And Mod(记录状态,3)=1 And 审核人 Is Null And (库房ID+0=[3] Or 库房ID Is NULL)"
    Else
        gstrSQL = "" & _
        "   Select NO,药品ID,批次,零售价,实际数量,付数,零售金额 " & _
        "   From 药品收发记录 " & _
        "   Where No=[1] And 单据=[2]" & _
        "            And Mod(记录状态,3)=1 And 审核人 Is Null And (库房ID+0=[3] Or 库房ID Is NULL)"
    End If
    
    '显示汇总数据
    Select Case mintUnit
    Case 0
        str明细单位串 = "C.计算单位 单位,B.零售价 单价,B.实际数量*Nvl(B.付数,1) 数量"
        str汇总单位串 = "C.计算单位 单位,B.零售价 单价,Sum(B.实际数量*Nvl(B.付数,1)) 数量"
    Case Else
        str明细单位串 = "D.包装单位 单位,B.零售价*nvl(D.换算系数,1) 单价,B.实际数量/nvl(D.换算系数,1)*Nvl(B.付数,1) 数量"
        str汇总单位串 = "D.包装单位 单位,B.零售价*nvl(D.换算系数,1) 单价,Sum(B.实际数量/nvl(D.换算系数,1)*Nvl(B.付数,1)) 数量"
    End Select
    
    str明细单位串 = str明细单位串 & ",B.零售金额 金额 "
    str汇总单位串 = str汇总单位串 & ",Sum(B.零售金额) 金额 "
    
    
    gstrSQL = "Select Distinct D.*,'['||D.编码||']'||D.通用名称  As 品名" & _
             " From (   SELECT B.NO,D.材料ID,C.编码,C.名称 通用名称,NVL(B.批次,0) 批次," & _
             "                  DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地)) 规格," & str汇总单位串 & _
             "          FROM (" & gstrSQL & ") B, 材料特性 D,收费项目目录 C " & _
             "          WHERE B.药品ID+0=D.材料ID AND D.材料ID=C.ID" & _
             "          GROUP BY B.NO,D.材料ID,C.编码,C.名称,NVL(B.批次,0)," & _
             "                 DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地)),"
    
    Select Case mintUnit
    Case 0
        gstrSQL = gstrSQL & "C.计算单位,B.零售价"
    Case Else
        gstrSQL = gstrSQL & "D.包装单位,B.零售价*nvl(D.换算系数,1)"
    End Select
    gstrSQL = gstrSQL & ") D"
    gstrSQL = gstrSQL & " Order By D.编码"
    
    err = 0: On Error Resume Next
    Set mrsTotal = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr单据号, mint单据类型, mlng发料部门ID)
    
    If mbln按票据号发料 Then
        '删除以前的单据
        With mrs待发汇总明细
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                If Not .EOF Then .Delete
                If Not .EOF Then .MoveNext
            Loop
        End With
    End If
    Call WriteTotalDataToBill
    
    If err <> 0 Then
        MsgBox "显示汇总数据时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnModify = False
    RefreshData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteTotalDataToBill() As Boolean
    Dim Dbl金额 As Double
    '将汇总数据装入
    On Error Resume Next
    err = 0
    WriteTotalDataToBill = False
    With Msf待发汇总
        .Clear
        .Rows = 2
        Call SetFormat(3)
    End With
    
    '填充单据内容
    Dbl金额 = 0
    
    If mrsTotal.State = 0 Then Exit Function
    
    If mrsTotal.RecordCount > 0 Then
        Do While Not mrsTotal.EOF
            With mrs待发汇总明细
                .AddNew
                !单据号 = mrsTotal!NO
                !材料名称 = mrsTotal!品名
                !编码 = mrsTotal!编码
                !规格 = IIf(IsNull(mrsTotal!规格), "", mrsTotal!规格)
                !单位 = IIf(IsNull(mrsTotal!单位), "", mrsTotal!单位)
                !单价 = mrsTotal!单价
                !数量 = mrsTotal!数量
                !金额 = mrsTotal!金额
                !材料ID = mrsTotal!材料ID
                !批次 = mrsTotal!批次
            End With
            mrsTotal.MoveNext
        Loop
    End If
    
    With mrs待发汇总明细
        If .RecordCount <> 0 Then
            .Sort = "编码,批次"
            .MoveFirst
        End If
        Do While Not .EOF
            If Msf待发汇总.Rows = 2 And Msf待发汇总.TextMatrix(1, 1) = "" Then
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 0) = Msf待发汇总.Rows - 1
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 1) = !材料名称
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 2) = IIf(IsNull(!规格), "", !规格)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 3) = IIf(IsNull(!单位), "", !单位)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 4) = Format(!单价, mFMT.FM_零售价)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 5) = Format(!数量, mFMT.FM_数量)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 6) = Format(!金额, mFMT.FM_金额)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 7) = !材料ID
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 8) = !批次
                Msf待发汇总.MergeRow(Msf待发汇总.Rows - 1) = False
            ElseIf Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 7) <> !材料ID Then
                Msf待发汇总.Rows = Msf待发汇总.Rows + 1
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 0) = Msf待发汇总.Rows - 1
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 1) = !材料名称
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 2) = IIf(IsNull(!规格), "", !规格)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 3) = IIf(IsNull(!单位), "", !单位)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 4) = Format(!单价, mFMT.FM_零售价)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 5) = Format(!数量, mFMT.FM_数量)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 6) = Format(!金额, mFMT.FM_金额)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 7) = !材料ID
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 8) = !批次
                Msf待发汇总.MergeRow(Msf待发汇总.Rows - 1) = False
            ElseIf Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 8) <> !批次 Then
                Msf待发汇总.Rows = Msf待发汇总.Rows + 1
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 0) = Msf待发汇总.Rows - 1
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 1) = !材料名称
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 2) = IIf(IsNull(!规格), "", !规格)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 3) = IIf(IsNull(!单位), "", !单位)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 4) = Format(!单价, mFMT.FM_零售价)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 5) = Format(!数量, mFMT.FM_数量)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 6) = Format(!金额, mFMT.FM_金额)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 7) = !材料ID
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 8) = !批次
                Msf待发汇总.MergeRow(Msf待发汇总.Rows - 1) = False
            Else
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 5) = Format(CDbl(IIf(Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 5) = "", 0, Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 5))) + !数量, mFMT.FM_数量)
                Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 6) = Format(CDbl(IIf(Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 6) = "", 0, Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 6))) + !金额, mFMT.FM_金额)
            End If
            Dbl金额 = Dbl金额 + !金额
            .MoveNext
        Loop
        
        '显示合计
        Msf待发汇总.Rows = Msf待发汇总.Rows + 1
        Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 0) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 1) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 2) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 3) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 4) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 5) = Format(Dbl金额, mFMT.FM_金额)
        Msf待发汇总.TextMatrix(Msf待发汇总.Rows - 1, 6) = Format(Dbl金额, mFMT.FM_金额)
        Msf待发汇总.MergeCells = flexMergeFree
        Msf待发汇总.MergeRow(Msf待发汇总.Rows - 1) = True
    End With
    
    If err <> 0 Then
        MsgBox "显示单据时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    WriteTotalDataToBill = True
End Function

Private Function WriteDataToBill() As Boolean
    Dim dbl合计金额 As Double
    '--显示指定处方的明细--
    On Error Resume Next
    err = 0
    
    WriteDataToBill = False
    With Msf待发明细
        .Clear
        .Rows = 2
        Call SetFormat(2)
    End With
    dbl合计金额 = 0
    
    '填充单据内容
    With mrsBill
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Msf待发明细.MergeRow(.AbsolutePosition) = False
            Msf待发明细.TextMatrix(.AbsolutePosition, 0) = !品名
            Msf待发明细.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!规格), "", !规格)
            Msf待发明细.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!单位), "", !单位)
            Msf待发明细.TextMatrix(.AbsolutePosition, 3) = Format(!单价, mFMT.FM_零售价)
            Msf待发明细.TextMatrix(.AbsolutePosition, 4) = Format(!数量, mFMT.FM_数量)
            Msf待发明细.TextMatrix(.AbsolutePosition, 5) = Format(!金额, mFMT.FM_金额)
            dbl合计金额 = dbl合计金额 + Val(!金额)
            
            If .AbsolutePosition >= Msf待发明细.Rows - 1 Then Msf待发明细.Rows = Msf待发明细.Rows + 1
            .MoveNext
        Loop
    End With
    With Msf待发明细
        .TextMatrix(.Rows - 1, 0) = "合计"
        .TextMatrix(.Rows - 1, 1) = "合计"
        .TextMatrix(.Rows - 1, 2) = "合计"
        .TextMatrix(.Rows - 1, 3) = "合计"
        .TextMatrix(.Rows - 1, 4) = Format(dbl合计金额, mFMT.FM_金额)
        .TextMatrix(.Rows - 1, 5) = Format(dbl合计金额, mFMT.FM_金额)
        .MergeCells = flexMergeFree
        .MergeRow(.Rows - 1) = True
    End With
    
    If err <> 0 Then
        MsgBox "显示单据时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    WriteDataToBill = True
End Function

Private Function SetLocateBill(ByVal strNo As String, ByVal intBillType As Integer, Optional ByVal BlnEnterCell As Boolean = True) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:找到指定单据是否存在
    '参数:strNo-单据号
    '     intBillType-单据类型
    '     BlnEnterCell-是否点击待发列表
    '返回:找到了返回true,否则返回false
    Dim intRow As Integer
    
    SetLocateBill = False
    With Msf待发列表
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 1) = strNo And intBillType = .RowData(intRow) Then
                .Row = intRow
                .TopRow = intRow
                SetLocateBill = True
                Exit For
            End If
        Next
    End With
    
    If BlnEnterCell Then Msf待发列表_EnterCell
End Function

Private Function CheckStock() As Boolean
    Dim rsCheckStock As New ADODB.Recordset
    Dim dblStock As Double
    Dim strSubSql As String
    Dim n As Integer
    
    '检查库存
    On Error GoTo ErrHandle
    If mIntCheckStock = 0 Then CheckStock = True: Exit Function
    
    '将库存数量转换为对应单位的实际数量
    Select Case mintUnit
    Case 0
        strSubSql = "/1"
    Case Else
        strSubSql = "/Decode(B.换算系数,0,1,null,1,b.换算系数)"
    End Select
    
    CheckStock = False
    If Msf待发列表.TextMatrix(1, 1) <> "" Then
        For n = 1 To Msf待发汇总.Rows - 2
            
            gstrSQL = "" & _
                "   Select nvl(实际数量,0)" & strSubSql & " AS 数量" & _
                "   From 药品库存 A,材料特性 B" & _
                "   Where A.药品ID=B.材料ID And A.性质=1 And A.库房ID=[3]" & _
                "           And A.药品ID=[1] And Nvl(A.批次,0)=[2]"
        
            Set rsCheckStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Msf待发汇总.TextMatrix(n, 7)), Val(Msf待发汇总.TextMatrix(n, 8)), mlng发料部门ID)
            With rsCheckStock
                If .EOF Then
                    dblStock = 0
                Else
                    dblStock = !数量
                End If
                
                If dblStock < Msf待发汇总.TextMatrix(n, 5) Then
                    If Msf待发汇总.TextMatrix(n, 8) <> 0 Then
                        MsgBox Msf待发汇总.TextMatrix(n, 1) & "的批次库存数不够，不能继续发料！", vbInformation, gstrSysName: Exit Function
                    Else
                        Select Case mIntCheckStock
                        Case 1
                            If MsgBox(Msf待发汇总.TextMatrix(n, 1) & "的库存数不够，是否继续发料？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        Case 2
                            MsgBox Msf待发汇总.TextMatrix(n, 1) & "的库存数不够，不能继续发料！", vbInformation, gstrSysName: Exit Function
                        End Select
                    End If
                End If
            End With
        Next
    End If
    CheckStock = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SendBill() As Boolean
    Dim intRow As Integer
    Dim strDate As String
    Dim strNo As String
    Dim str单据 As String
    Dim lngCount As Long
    Dim strNos As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    err = 0
    SendBill = False
    
    strDate = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
    lngCount = 0
    gcnOracle.BeginTrans
    With Msf待发列表
        For intRow = 1 To .Rows - 1
            If .RowData(intRow) <> 0 Then
                '检查处方
                If CheckBill(.RowData(intRow), .TextMatrix(intRow, 1)) <> 0 Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
                strNo = Trim(.TextMatrix(intRow, 1))
                str单据 = .RowData(intRow)
                '----:发料方式_IN：单据发药为1；批量发药为2；部门发药为3
                '过程参数:库房ID_IN,单据_IN,NO_IN,审核人_IN,配料人_IN,校验人_IN,发料方式_IN,审核日期_IN
                gstrSQL = "zl_材料收发记录_处方发料(" & _
                    mlng发料部门ID & "," & _
                    .RowData(intRow) & ",'" & _
                    .TextMatrix(intRow, 1) & "','" & _
                    gstrUserName & "','" & _
                    gstrUserName & "',NULL," & _
                    2 & ",to_date('" & _
                    strDate & "','yyyy-MM-dd hh24:mi:ss'))"
                
                strNos = IIf(strNos = "", "", strNos & "|") & .RowData(intRow) & "," & .TextMatrix(intRow, 1)
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-材料发料")
               lngCount = lngCount + 1
            End If
        Next
    End With
    gcnOracle.CommitTrans
    
    If lngCount = 0 Then
    Else
        If lngCount = 1 Then
            Call BillListPrint(1, strDate, strNo, str单据)
        Else
            Call BillListPrint(2, strDate)
        End If
    End If
    
    '调用发料后的外挂接口
    If Not mobjPlugIn Is Nothing And strNos <> "" Then
        mobjPlugIn.StuffSendByRecipe mlng发料部门ID, strNos, CDate(strDate), strReserve
    End If
    
    SendBill = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub BillListPrint(Optional int发料方式 As Integer = 1, Optional strDate As String = "", Optional strNo As String = "", Optional str单据 As String = "0")
    '单据或清册打印
    '发料方式:1-处方发料;2-批量发料;3-部门发料
    ' intStyle:0-按发料方式打印,1-单据打印
    Dim bln退料单 As Boolean
    Dim bln已发料清单 As Boolean
    Dim bln单据打印 As Boolean
    Dim strReg As String
    Dim intPrint As Integer '0-提示打印,1-自动打印,<>0或1:不打印
    
    bln退料单 = InStr(1, mstrPrivs, "退料通知单") <> 0
    bln已发料清单 = InStr(1, gstrPrivs, "打印已发料清单") <> 0
    bln单据打印 = InStr(1, gstrPrivs, "单据打印") <> 0
    
    If bln单据打印 = False Then Exit Sub
    
    intPrint = Val(zlDatabase.GetPara("发料打印提醒方式", glngSys, mlngModule, "0"))
    
    If intPrint = 0 Then
        '提示打印
        If MsgBox("你需要打印相关单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    ElseIf intPrint = 1 Then
        '自动打印
    Else
        Exit Sub
    End If
    Select Case int发料方式
    Case 1  '处方打印
        If strNo <> "" Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723", Me, "库房==" & mlng发料部门ID, "NO=" & strNo, "单据=" & str单据, "审核人=审核人 is not null", 1)
        Else
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "库房=" & mlng发料部门ID, "发料方式=单据发料|1", "发料号=" & strDate, 1)
        End If
    Case 2
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "库房=" & mlng发料部门ID, "发料方式=批量发料|2", "发料号=" & strDate, 1)
    Case 3
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "库房=" & mlng发料部门ID, "发料方式=部门发料|3", "发料号=" & strDate, 1)
    End Select
    
End Sub

Private Function CheckCorrelation() As Boolean
    Dim strNo As String, lng单据 As Long, str序号 As String
    '检查处方是否已结帐、检查该病人是否已出院，并对权限进行检查
    '暂无此方面的检查
'    With mrs序号
'        If .RecordCount <> 0 Then .MoveFirst
'        Do While Not .EOF
'            StrNo = !单据标识
'            lng单据 = Split(StrNo, "|")(1)
'            StrNo = Split(StrNo, "|")(0)
'            str序号 = NVL(!序号)
'            '暂无“发退结帐”处方的权限，因此屏蔽
'            'If Not IsReceiptBalance_Charge(mstrPrivs, lng单据, StrNo, str序号) Then Exit Function
'            '检查出院病人
'            If Not IsOutPatient(mstrPrivs, lng单据, StrNo) Then Exit Function
'            .MoveNext
'        Loop
'    End With
'
    CheckCorrelation = True
End Function

Private Sub InitRec()
    Set mrs序号 = New ADODB.Recordset
    With mrs序号
        If .State = 1 Then .Close
        .Fields.Append "单据标识", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "序号", adLongVarChar, 500, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set mrs处方来源部门 = New ADODB.Recordset
    With mrs处方来源部门
        If .State = 1 Then .Close
        .Fields.Append "编码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "来源部门", adLongVarChar, 100, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set mrs待发汇总明细 = New ADODB.Recordset
    With mrs待发汇总明细
        If .State = 1 Then .Close
        .Fields.Append "单据号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "材料名称", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "编码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "单价", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adDouble, 18, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
