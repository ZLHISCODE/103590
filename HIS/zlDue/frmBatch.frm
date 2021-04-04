VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "批量付款计划"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12030
   Icon            =   "frmBatch.frx":0000
   LinkTopic       =   "批量付款计划"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk全选 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "全选"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   7800
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6480
      Width           =   675
   End
   Begin VB.CommandButton cmd全部付款 
      Caption         =   "全部付款(&A)"
      Height          =   300
      Left            =   10680
      TabIndex        =   29
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frmline2 
      Height          =   120
      Left            =   120
      TabIndex        =   24
      Top             =   6120
      Width           =   11895
   End
   Begin VB.CommandButton cmd批量分配 
      Caption         =   "批量分配(&B)"
      Height          =   300
      Left            =   9240
      TabIndex        =   22
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Width           =   1100
   End
   Begin VB.Frame Frmline1 
      Height          =   120
      Left            =   120
      TabIndex        =   14
      Top             =   620
      Width           =   11895
   End
   Begin VB.TextBox txt计划金额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7440
      TabIndex        =   7
      Top             =   960
      Width           =   1710
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   9600
      TabIndex        =   10
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10800
      TabIndex        =   9
      Top             =   6360
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtp付款时间 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   4440
      TabIndex        =   11
      Top             =   960
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   124125187
      CurrentDate     =   36846.5833333333
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   4725
      Left            =   3480
      TabIndex        =   23
      Top             =   1350
      Width           =   8460
      _cx             =   14922
      _cy             =   8334
      Appearance      =   0
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBatch.frx":6852
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
      ExplorerBar     =   5
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
   Begin VB.Frame fraTemp 
      Caption         =   "提取单据条件"
      Height          =   5235
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   3345
      Begin VB.TextBox txt供应商 
         Height          =   300
         Left            =   1140
         TabIndex        =   1
         Top             =   480
         Width           =   1770
      End
      Begin VB.ComboBox cbo审核日期 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   2085
      End
      Begin VB.OptionButton optClass 
         Caption         =   "药品(&1)"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Value           =   -1  'True
         Width           =   1000
      End
      Begin VB.OptionButton optClass 
         Caption         =   "卫材(&2)"
         Height          =   180
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         Top             =   2400
         Width           =   1000
      End
      Begin VB.CommandButton cmd提取单据 
         Caption         =   "提取单据"
         Height          =   350
         Left            =   2125
         TabIndex        =   19
         Top             =   3000
         Width           =   1100
      End
      Begin VB.CommandButton cmd供应商 
         Caption         =   "…"
         Height          =   300
         Left            =   2880
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   465
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   1410
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   124125187
         CurrentDate     =   40848
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   1905
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   148504579
         CurrentDate     =   40848
      End
      Begin VB.Label lbl供应商 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "供 应 商"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lbl结束日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   1965
         Width           =   720
      End
      Begin VB.Label lbl开始日期 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Label lbl未付金额数 
      AutoSize        =   -1  'True
      Caption         =   "30000000"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6120
      TabIndex        =   21
      Top             =   6480
      Width           =   840
   End
   Begin VB.Label lbl总金额数 
      AutoSize        =   -1  'True
      Caption         =   "50000000"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4080
      TabIndex        =   20
      Top             =   6480
      Width           =   840
   End
   Begin VB.Label lblInfor 
      Caption         =   "根据外购入库单已审核的单据，批量制定付款计划。"
      Height          =   285
      Left            =   810
      TabIndex        =   15
      Top             =   360
      Width           =   8535
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   255
      Picture         =   "frmBatch.frx":695D
      Top             =   100
      Width           =   480
   End
   Begin VB.Label lbl未付金额 
      AutoSize        =   -1  'True
      Caption         =   "未付金额："
      Height          =   180
      Left            =   5280
      TabIndex        =   13
      Top             =   6495
      Width           =   900
   End
   Begin VB.Label lbl计划金额 
      AutoSize        =   -1  'True
      Caption         =   "计划金额"
      Height          =   180
      Left            =   6600
      TabIndex        =   8
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lbl总金额 
      AutoSize        =   -1  'True
      Caption         =   "总金额："
      Height          =   180
      Left            =   3360
      TabIndex        =   0
      Top             =   6495
      Width           =   720
   End
   Begin VB.Label lbl付款时间 
      BackColor       =   &H80000004&
      Caption         =   "付款时间"
      Height          =   180
      Left            =   3480
      TabIndex        =   12
      Top             =   1005
      Width           =   855
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlng供应商ID As Long
Private mint状态 As Integer  '0--增加；1--删除
Private mstr供应商 As String
Private mblnOK As Boolean
Private Const mconlngColor As Long = &HFFFFFF        '不能修改列颜色为白色
Private Const mconlngCanColColor As Long = &HFFE3C8        '能修改列颜色为淡蓝色
Private Const mlngBorderColor As Long = &H0&    '选中行边框颜色
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' 没选中行边框颜色

Public Sub ShowCard(ByVal frmMain As Object, ByVal strPrivs As String, Optional lng供应商ID As Long, Optional str供应商 As String, Optional int状态 As Integer)

    mstrPrivs = strPrivs
    mstr供应商 = str供应商
    mlng供应商ID = lng供应商ID
    mint状态 = int状态
    mblnOK = False
    Me.Show vbModal, frmMain
End Sub

Private Function ValidData() As Boolean
    Dim i As Integer
    Dim dbl计划金额 As Double
    Dim bln删除 As Boolean
    
    If vsfList.Rows < 2 Then Exit Function
    With vsfList
        If mint状态 = 0 Then
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("计划金额"))) > Val(.TextMatrix(i, .ColIndex("应付金额"))) Then
                    MsgBox "第" & i & "行计划金额大于了应付金额，请检查！", vbOKOnly + vbInformation, gstrSysName
                    .Row = i
                    .Col = .ColIndex("计划金额")
                    .TopRow = i
                    Exit Function
                End If
                
                dbl计划金额 = dbl计划金额 + Val(.TextMatrix(i, .ColIndex("计划金额")))
            Next
            
            If dbl计划金额 = 0 Then
                MsgBox "所有计划金额都为0或空了，请检查！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If mint状态 = 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("行号")) = "√" Then
                    bln删除 = True
                    Exit For
                End If
            Next
            
            If bln删除 = False Then
                MsgBox "请先勾选一个单据后再保存！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    ValidData = True
End Function

Private Sub chk全选_Click()
    Dim i As Integer
    
    If vsfList.Rows < 2 Then chk全选.Value = 0: Exit Sub
    With vsfList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("行号")) = "" And chk全选.Value = 1 Then
                .TextMatrix(i, .ColIndex("行号")) = "√"
            ElseIf .TextMatrix(i, .ColIndex("行号")) = "√" And chk全选.Value = 0 Then
                .TextMatrix(i, .ColIndex("行号")) = ""
            End If
        Next
    End With
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim arrSql() As Variant     '纪录存储过程的数组
    Dim blnTrans As Boolean
    Dim i As Integer
    Dim dbl计划金额 As Double
    
    On Error GoTo ErrHand:
    
    If ValidData = False Then Exit Sub

    If Format(dtp付款时间.Value, "yyyy-MM-dd") < Format(Sys.Currentdate, "yyyy-MM-dd") Then
        If MsgBox("计划付款日期小于制定计划日期，是否确定保存?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox(IIf(mint状态 = 0, "是否确定保存？", "是否确定删除？"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
        
    arrSql = Array()
    
    For i = 1 To vsfList.Rows - 1
        If mint状态 = 0 Then
            If Val(vsfList.TextMatrix(i, vsfList.ColIndex("计划金额"))) > 0 Then
                dbl计划金额 = Val(vsfList.TextMatrix(i, vsfList.ColIndex("计划金额")))
                
                gstrSQL = "Select ID, 计划序号, 剩余应付金额" & vbNewLine & _
                                "From (Select a.Id, Max(Nvl(a.计划序号, 0)) As 计划序号, Sum(Nvl(a.发票金额, 0)) / Count(1) - Sum(Nvl(a.计划金额, 0)) As 剩余应付金额" & vbNewLine & _
                                "       From 应付记录 A, 药品收发记录 B" & vbNewLine & _
                                "       Where a.收发id = b.Id And b.单据 = [2]  And b.No =[1] And b.审核日期 Is Not Null" & vbNewLine & _
                                "       Group By a.Id" & vbNewLine & _
                                "       Order By a.Id)" & vbNewLine & _
                                "Where 剩余应付金额 > 0"
    
                Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "批量提取药品明细", vsfList.TextMatrix(i, vsfList.ColIndex("单据号")), IIf(optClass(0).Value = True, 1, 15))
        
                With rsTmp
                    Do While Not .EOF
                    
                        If dbl计划金额 > 0 Then
                            gstrSQL = "ZL_付款计划_INSERT("
                            
                            'ID_IN        IN 应付记录.ID%Type,
                            gstrSQL = gstrSQL & !ID
                            
                            '计划序号_IN    IN 应付记录.计划序号%Type,
                            gstrSQL = gstrSQL & "," & !计划序号 + 1
                            
                            '计划金额_IN    IN 应付记录.计划金额%Type,
                            gstrSQL = gstrSQL & "," & IIf(dbl计划金额 > !剩余应付金额, !剩余应付金额, dbl计划金额)
                            
                            '计划日期_IN    IN 应付记录.计划日期%Type,
                            gstrSQL = gstrSQL & "," & "TO_DATE('" & Format(dtp付款时间.Value, "yyyy-MM-dd") & "','yyyy-MM-dd')"
                            
                            '计划人_IN    IN 应付记录.计划人%Type,
                            gstrSQL = gstrSQL & ",'" & UserInfo.姓名 & "'"
                            
                            '制定日期_IN    IN 应付记录.制定日期%Type
                            gstrSQL = gstrSQL & "," & "TO_DATE('" & Format(Sys.Currentdate, "yyyy-MM-dd") & "','yyyy-MM-dd')"
                            
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                            arrSql(UBound(arrSql)) = gstrSQL
                            
                            dbl计划金额 = dbl计划金额 - !剩余应付金额
                        End If
                        
                        .MoveNext
                    Loop
                End With
            End If
        Else
            If vsfList.TextMatrix(i, vsfList.ColIndex("行号")) = "√" Then
                gstrSQL = "Select a.id" & vbNewLine & _
                                "From 应付记录 A, 药品收发记录 B" & vbNewLine & _
                                "Where a.收发id = b.Id And b.单据 = [3] And b.审核日期 Is Not Null And a.记录性质 = -1 And a.付款序号 Is Null and b.no=[1]" & vbNewLine & _
                                "And a.计划日期 =[2]"

                Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "批量提取药品明细", vsfList.TextMatrix(i, vsfList.ColIndex("单据号")), _
                CDate(Format(vsfList.TextMatrix(i, vsfList.ColIndex("单据日期")), "yyyy-mm-dd")), IIf(optClass(0).Value = True, 1, 15))
        
                With rsTmp
                    Do While Not .EOF
                        gstrSQL = "ZL_付款计划_DELETE("
                        
                        'ID_IN        IN 应付记录.ID%Type,
                        gstrSQL = gstrSQL & !ID
                        
                        '计划日期_In    IN 应付记录.开始日期_In,
                        gstrSQL = gstrSQL & "," & "TO_DATE('" & Format(vsfList.TextMatrix(i, vsfList.ColIndex("单据日期")), "yyyy-MM-dd") & "','yyyy-MM-dd')"
                        
                        gstrSQL = gstrSQL & ")"
                        
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = gstrSQL
                        
                        .MoveNext
                    Loop
                End With
            End If
        End If
    Next
                
    gcnOracle.BeginTrans: blnTrans = True          '开启事务
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False     '提交事物
    
    mblnOK = True
    MsgBox IIf(mint状态 = 0, "保存成功！", "删除成功！"), vbOKOnly + vbInformation, gstrSysName
    cmd提取单据_Click
    mblnOK = False
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd批量分配_Click()
    Dim i As Integer
    Dim dbl计划金额 As Double
    
    If vsfList.Rows < 2 Then
        Exit Sub
    End If
    
    If Val(txt计划金额.Text) > Val(lbl总金额数.Caption) Then
        MsgBox "计划金额大于了总金额 [" & Format(Val(lbl总金额数.Caption), "0.00") & "]，请重新输入！", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If

    If Val(txt计划金额.Text) = 0 Then
        MsgBox "计划金额不能为空或0，请检查！", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    
    dbl计划金额 = Val(txt计划金额.Text)
    With vsfList
        For i = 1 To .Rows - 1
        
            If dbl计划金额 > Val(.TextMatrix(i, .ColIndex("应付金额"))) Then
                .TextMatrix(i, .ColIndex("计划金额")) = Format(.TextMatrix(i, .ColIndex("应付金额")), "0.00")
            Else
                .TextMatrix(i, .ColIndex("计划金额")) = Format(dbl计划金额, "0.00")
            End If

            dbl计划金额 = dbl计划金额 - Val(.TextMatrix(i, .ColIndex("计划金额")))
        Next
    End With
    
    
    lbl未付金额数.Caption = Format(Val(lbl总金额数.Caption) - Val(txt计划金额.Text), "0.00")
    
End Sub

Private Sub cmd全部付款_Click()
    txt计划金额.Text = lbl总金额数.Caption
    cmd批量分配_Click
End Sub

Private Sub cmd提取单据_Click()
    Dim rsTemp As ADODB.Recordset
    Dim dbl总金额 As Double
    
    On Error GoTo ErrHand:
    
    vsfList.Rows = 1
    
    If mint状态 = 0 Then
            gstrSQL = "Select NO, 审核日期, 摘要, Sum(Nvl(付款金额, 0)) - Sum(Nvl(计划金额, 0)) As 应付金额" & vbNewLine & _
                            "From (Select b.记录性质, a.No, a.审核日期, a.摘要, Decode(记录性质, 0, b.发票金额, 0) As 付款金额, Decode(记录性质, -1, b.计划金额, 0) As 计划金额" & vbNewLine & _
                            "       From 药品收发记录 A, 应付记录 B, " & IIf(optClass(0).Value = True, "药品规格 D", "材料特性 D") & vbNewLine & _
                            "       Where a.Id = b.收发id And a.供药单位id + 0 = [1] And" & vbNewLine & _
                            "             a.审核日期 Between [2] And [3] And a.单据 =[4] And a.药品id = " & IIf(optClass(0).Value = True, "d.药品id", "d.材料id") & ")" & vbNewLine & _
                            "Group By NO, 审核日期, 摘要" & vbNewLine & _
                            "Having Sum(付款金额) - Sum(计划金额) > 0" & vbNewLine & _
                            "Order By NO, 审核日期"

            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "批量提取付款计划", Val(txt供应商.Tag), _
            CDate(Format(dtp开始时间.Value, "yyyy-mm-dd")), CDate(Format(dtp结束时间.Value, "yyyy-mm-dd") & " 23:59:59"), _
             IIf(optClass(0).Value = True, 1, 15))
            
            If rsTemp.RecordCount = 0 And mblnOK = False Then
                MsgBox "未查询到符合条件的单据，请检查！", vbOKOnly + vbInformation, gstrSysName
            End If
            
            With rsTemp
                Do While Not .EOF
                    vsfList.Rows = vsfList.Rows + 1
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("行号")) = vsfList.Rows - 1
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("单据号")) = !No
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("单据日期")) = Format(!审核日期, "yyyy-mm-dd hh:mm:ss")
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("摘要")) = Nvl(!摘要)
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("计划金额")) = Format(0, "0.00")
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("应付金额")) = Format(!应付金额, "0.00")
                    
                    dbl总金额 = dbl总金额 + !应付金额
                rsTemp.MoveNext
                Loop
            End With
            
            lbl总金额数.Caption = Format(dbl总金额, "0.00")
            lbl未付金额数.Caption = Format(dbl总金额, "0.00")
            txt计划金额.Text = Format(0, "0.00")
            
            Call setColEdit
        Else
            gstrSQL = "Select b.No, a.计划日期, b.摘要, Sum(Nvl(a.计划金额, 0)) As 计划金额, Sum(Nvl(a.发票金额, 0)) As 应付金额" & vbNewLine & _
                            "From 应付记录 A, 药品收发记录 B, " & IIf(optClass(0).Value = True, "药品规格 D", "材料特性 D") & vbNewLine & _
                            "Where a.收发id = b.Id And b.药品id = " & IIf(optClass(0).Value = True, "d.药品id", "d.材料id") & " And b.供药单位id + 0 = [1] And b.单据 = [4] And a.记录性质 = -1" & vbNewLine & _
                            "      And a.付款序号 Is Null And a.计划日期 Between [2] And [3] And b.审核日期 Is Not Null" & vbNewLine & _
                            "Group By a.计划日期, b.No, b.摘要" & vbNewLine & _
                            "Order By b.No, a.计划日期"

            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "批量提取付款计划", Val(txt供应商.Tag), _
            CDate(Format(dtp开始时间.Value, "yyyy-mm-dd")), CDate(Format(dtp结束时间.Value, "yyyy-mm-dd") & " 23:59:59"), _
            IIf(optClass(0).Value = True, 1, 15))
            
            If rsTemp.RecordCount = 0 And mblnOK = False Then
                MsgBox "未查询到符合条件的单据，请检查！", vbOKOnly + vbInformation, gstrSysName
            End If
            
            With rsTemp
                Do While Not .EOF
                    vsfList.Rows = vsfList.Rows + 1
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("单据号")) = !No
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("单据日期")) = Format(!计划日期, "yyyy-mm-dd")
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("摘要")) = Nvl(!摘要)
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("计划金额")) = Format(!计划金额, "0.00")
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("应付金额")) = Format(!应付金额, "0.00")
                    
                rsTemp.MoveNext
                Loop
            End With
            
            chk全选.Value = 0
        End If
        
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call initComboBox
    Call Int初始化界面
End Sub


Private Sub txt供应商_Change()
    With txt供应商
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub txt供应商_GotFocus()
    txt供应商.SelStart = 0
    txt供应商.SelLength = Len(txt供应商.Text)
End Sub

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim adoProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim str权限 As String
    
    vRect = zlControl.GetControlRect(txt供应商.hwnd)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo errHandle
    With txt供应商
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        
        str权限 = " and " & Get分类权限(mstrPrivs)

        gstrSQL = "" & _
            "  Select   ID,编码,名称,简码,类型" & _
            "  From  供应商 " & _
            "  Where (撤档时间 is null or To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01') " & _
            "       " & zl_获取站点限制 & "  and 末级=1  " & _
            "       And ( 编码 Like [1] or 名称 like [1] or 简码  like upper([1])) " & str权限 & _
            "      order by  编码  "
            
        Set adoProvider = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "供药单位", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", strProviderText & "%", gstrNodeNo)

        If blnCancel = True Then .SetFocus: Exit Sub  '打开选择器时，点Esc不做以下处理
        
        If adoProvider.State = 0 Then
            MsgBox "没有你输入的供药单位，请重输！", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If

        .Text = "[" & adoProvider!编码 & "]" & adoProvider!名称
        .Tag = adoProvider!ID
        
        
        adoProvider.Close
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd供应商_Click()
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    strTemp = frm供应商选择.SelDept(mstrPrivs)
    If strTemp = "" Then
        Unload frm供应商选择
        If txt供应商.Enabled Then txt供应商.SetFocus
        Exit Sub
    End If
    txt供应商.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    txt供应商.Tag = Val(Left(strTemp, InStr(strTemp, ",") - 1))
    On Error GoTo errHandle
    Set rsTemp = zldatabase.OpenSQLRecord("select 类型 from 供应商 where id=[1] ", Caption & "-提取供应商类型", txt供应商.Tag)

    rsTemp.Close
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub initComboBox()
    With cbo审核日期
        .Clear
        .AddItem "今日"
        If mint状态 = 0 Then
            .AddItem "一星期内"
            .AddItem "一个月内"
            .AddItem "三个月内"
        Else
            .AddItem "未来一星期"
            .AddItem "未来一个月"
            .AddItem "未来三个月"
        End If
        .AddItem "自定义日期"
        .ListIndex = 0
    End With
End Sub

Private Sub cbo审核日期_Click()
    Dim dateCurrentDate As Date
    
    If cbo审核日期.Text = "自定义日期" Then
        dtp开始时间.Enabled = True
        dtp结束时间.Enabled = True
        
    Else
        dtp开始时间.Enabled = False
        dtp结束时间.Enabled = False
    End If
    
    '根据选择改变时间
    dateCurrentDate = Sys.Currentdate
    If mint状态 = 0 Then
        Select Case cbo审核日期.ListIndex
            Case 0
                dtp开始时间.Value = CDate(Format(dateCurrentDate, "yyyy-mm-dd") & " 00:00:00")
                dtp结束时间.Value = dateCurrentDate
            Case 1
                dtp开始时间.Value = CDate(Format(DateAdd("d", -7, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
                dtp结束时间.Value = dateCurrentDate
            Case 2
                dtp开始时间.Value = CDate(Format(DateAdd("d", -30, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
                dtp结束时间.Value = dateCurrentDate
            Case 3
                dtp开始时间.Value = CDate(Format(DateAdd("d", -90, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
                dtp结束时间.Value = dateCurrentDate
        End Select
    Else
        Select Case cbo审核日期.ListIndex
            Case 0
                dtp开始时间.Value = dateCurrentDate
                dtp结束时间.Value = CDate(Format(dateCurrentDate, "yyyy-mm-dd") & " 00:00:00")
            Case 1
                dtp开始时间.Value = dateCurrentDate
                dtp结束时间.Value = CDate(Format(DateAdd("d", 7, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            Case 2
                dtp开始时间.Value = dateCurrentDate
                dtp结束时间.Value = CDate(Format(DateAdd("d", 30, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            Case 3
                dtp开始时间.Value = dateCurrentDate
                dtp结束时间.Value = CDate(Format(DateAdd("d", 90, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
        End Select
    End If
End Sub

Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt供应商, KeyAscii, m文本式
End Sub

Private Sub txt计划金额_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt计划金额, KeyAscii, m金额式
End Sub

Private Sub txt计划金额_LostFocus()
    txt计划金额.Text = Format(txt计划金额.Text, "0.00")
End Sub

Private Sub vsfList_DblClick()
    If mint状态 = 1 Then
        With vsfList
            If .Rows < 2 Then Exit Sub
            .TextMatrix(.Row, .ColIndex("行号")) = IIf(.TextMatrix(.Row, .ColIndex("行号")) = "", "√", "")
        End With
    End If
End Sub

Private Sub vsfList_EnterCell()
    Dim i As Integer
    With vsfList
        If .Row = 0 Then Exit Sub
        .Editable = flexEDNone
        .FocusRect = flexFocusLight
        
        If mint状态 = 0 Then
            If .Row > 0 And .Col = .ColIndex("计划金额") Then
                .Editable = flexEDKbdMouse
                .FocusRect = flexFocusSolid
            End If
    
            If .Rows > 1 Then
                For i = 1 To .Rows - 1
                    .CellBorderRange i, 0, i, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
                Next
                .CellBorderRange .Row, 0, .Row, .Cols - 1, mlngBorderColor, 0, 2, 0, 2, 0, 2
            End If
            
            vsfList.Col = vsfList.ColIndex("计划金额")
        End If
    End With
    
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then Exit Sub
    If vsfList.Row = 0 Then Exit Sub
    With vsfList
        If .Rows > 1 Then
            If MsgBox("是否确定删除第" & .Row & "行单据号为“" & .TextMatrix(.Row, .ColIndex("单据号")) & "”的记录吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            Else
                .RemoveItem .Row
            End If
        End If
    End With
    
    Call 金额计算
End Sub

Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    With vsfList
        If KeyAscii = vbKeyBack Then Exit Sub
        Select Case .Col
            Case .ColIndex("计划金额")

                VsFlxGridCheckKeyPress vsfList, Row, Col, KeyAscii, m金额式

                strKey = .EditText
                If strKey = "" Then
                    strKey = .TextMatrix(.Row, .Col)
                End If
                
                If LenB(StrConv(.EditText, vbFromUnicode)) >= 16 Then
                    KeyAscii = 0
                End If
                                 
        End Select
    End With
End Sub


Private Sub 金额计算()
    Dim dbl总金额 As Double
    Dim dbl计划金额 As Double
    Dim i As Integer
    
    If vsfList.Rows < 2 Then Exit Sub
    With vsfList
        For i = 1 To .Rows - 1
            dbl总金额 = dbl总金额 + Val(.TextMatrix(i, .ColIndex("应付金额")))
            dbl计划金额 = dbl计划金额 + Val(.TextMatrix(i, .ColIndex("计划金额")))
        Next
    End With
    
    lbl总金额数.Caption = Format(dbl总金额, "0.00")
    lbl未付金额数.Caption = Format(dbl总金额 - dbl计划金额, "0.00")
    txt计划金额.Text = Format(dbl计划金额, "0.00")
    
End Sub


Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    If vsfList.Rows < 2 Then Exit Sub
    strKey = vsfList.EditText
    With vsfList
         .EditText = Format(Val(strKey), "0.00")
         .TextMatrix(.Row, .ColIndex("计划金额")) = Format(Val(strKey), "0.00")
        If Val(.TextMatrix(.Row, .ColIndex("计划金额"))) > Val(.TextMatrix(.Row, .ColIndex("应付金额"))) Then
            MsgBox "第" & .Row & "行计划金额大于了应付金额，请重新输入！", vbOKOnly + vbInformation, gstrSysName
        End If
    End With
    
    Call 金额计算
End Sub

Private Sub setColEdit()
    Dim intRow As Integer

    With vsfList
        If .Rows < 2 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = mconlngColor
        .Cell(flexcpBackColor, 1, .ColIndex("计划金额"), .Rows - 1, .ColIndex("计划金额")) = mconlngCanColColor
    End With

End Sub

Private Sub Int初始化界面()
    vsfList.Rows = 1
    txt供应商.Text = mstr供应商
    txt供应商.Tag = mlng供应商ID
    vsfList.AllowSelection = False '不能多选
    dtp付款时间.Value = Sys.Currentdate
    If mint状态 = 0 Then
        lbl总金额数.Caption = Format(0, "0.00")
        lbl未付金额数.Caption = Format(0, "0.00")
        txt计划金额.Text = Format(0, "0.00")
        chk全选.Visible = False
        frmBatch.Caption = "批量产生付款计划"
    Else
        lbl总金额.Visible = False
        lbl总金额数.Visible = False
        lbl未付金额.Visible = False
        lbl未付金额数.Visible = False
        lbl付款时间.Visible = False
        dtp付款时间.Visible = False
        lbl计划金额.Visible = False
        txt计划金额.Visible = False
        cmd批量分配.Visible = False
        cmd全部付款.Visible = False
        lbl审核日期.Caption = "付款日期"
        lblInfor.Caption = "根据外购入库单已审核的单据，批量删除表格中标记已勾选的付款计划。"
        chk全选.Left = vsfList.Left
        Frmline1.Top = 580
        vsfList.Top = 950
        vsfList.Height = 5110
        chk全选.Top = vsfList.Top - chk全选.Height - 10
        vsfList.TextMatrix(0, vsfList.ColIndex("行号")) = "标记"
        vsfList.TextMatrix(0, vsfList.ColIndex("单据日期")) = "付款日期"
        
        vsfList.ColWidth(vsfList.ColIndex("单据号")) = 1800
        vsfList.ColWidth(vsfList.ColIndex("单据日期")) = 2400
        
        vsfList.ColHidden(vsfList.ColIndex("应付金额")) = True
        frmBatch.Caption = "批量删除付款计划"
    End If
End Sub
