VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm卫材拒发料清单 
   BorderStyle     =   0  'None
   Caption         =   "卫材拒发料清单"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   4125
      Left            =   450
      TabIndex        =   0
      Top             =   210
      Width           =   7320
      _cx             =   12912
      _cy             =   7276
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm卫材拒发料清单.frx":0000
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
      WordWrap        =   -1  'True
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
Attribute VB_Name = "frm卫材拒发料清单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsNotPayStuff As ADODB.Recordset
Private mintUnit As Integer
Private mstrPrivs As String
Private mlngModule As Long
Private mArrFilter As Variant   '过滤条件
Private mfrmMain As Form        '父窗口
Private mrs拒发 As ADODB.Recordset
Private mlngCount恢复  As Long
Private Const mstrAllType As String = "临床,护理,检查,检验,手术,治疗,营养"
Private mbln按发生时间过滤 As Boolean

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Sub InitVsGrid()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始网格控件
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-12 10:27:06
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        '0-可选,1-必选,-1-隐藏
        .ColData(.ColIndex("状态")) = 1
        .ColData(.ColIndex("单据类型")) = 1
        .ColData(.ColIndex("单据号")) = 1
        .ColData(.ColIndex("材料名称")) = 1
        .ColData(.ColIndex("数量")) = 1
    End With
End Sub

Public Function zlRestorePayStuff() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:恢复已经拒发的卫生材料
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-24 11:40:50
    '-----------------------------------------------------------------------------------------------------------
    If ISValied = False Then Exit Function
    zlRestorePayStuff = SaveData()
End Function
Private Function ISValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查柜发的相关条件
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-24 11:41:49
    '-----------------------------------------------------------------------------------------------------------
    Dim blnHaveData As Boolean
    Dim lngRow As Long
    With vsGrid
        blnHaveData = False
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("状态")) = "恢复" Then
                blnHaveData = True: Exit For
            End If
        Next
        If blnHaveData = False Then
            ShowMsgBox "没有选择需要恢复的拒发材料，请操作中止!"
            Exit Function
        End If
    End With
    ISValied = True
End Function

Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:对卫生材料发放的拒发部分进行恢复操作
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-24 11:44:19
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, cllProc As Collection
    Dim int门诊 As Integer
    
    Set cllProc = New Collection
    With vsGrid
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("状态")) = "恢复" And .RowData(lngRow) <> 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("记录性质"))) = 1 Or (Val(.TextMatrix(lngRow, .ColIndex("记录性质"))) = 2 And (Val(.TextMatrix(lngRow, .ColIndex("门诊标志")))) = 1 Or (Val(.TextMatrix(lngRow, .ColIndex("门诊标志")))) = 4) Then
                    int门诊 = 1
                Else
                    int门诊 = 2
                End If
            
                'Zl_卫生材料发放_拒发恢复(Id_In In 药品收发记录.ID%Type)
                gstrSQL = "Zl_卫生材料发放_拒发恢复(" & .RowData(lngRow) & "," & int门诊 & ")"
                AddArray cllProc, gstrSQL
            End If
        Next
    End With
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllProc, Me.Caption
    mlngCount恢复 = 0
    SaveData = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlRefreshData(ByVal frmMain As Form, ByVal strPrivs As String, ByVal lngModule As Long, ByVal intUnit As Integer, _
    ByVal arrFilter As Variant) As Boolean
     '-----------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '入参:frmMain-父窗口
    '     strPrivs-权限串
    '     lngModule-模块号
    '     intUnit-显示单位(0-散装单位,1-包装单位)
    '     arrFilter-条件过滤
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-22 14:25:18
    '-----------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mstrPrivs = strPrivs: mlngModule = lngModule:
    Set mArrFilter = arrFilter
    mintUnit = intUnit
    
    '初始化值
    Call Form_Load
    With vsGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .Clear (1)
        '填充数据
        zlRefreshData = RefreshData
        .Redraw = flexRDBuffered
    End With
End Function
 
Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With vsGrid
        .Top = ScaleTop
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = ScaleHeight
    End With
End Sub
Public Function zlFullData(ByVal rsNotPayStuff As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:填充汇总数据到Vss控件中
    '入参:rsNotPayStuff-未发料清单
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 17:11:13
    '-----------------------------------------------------------------------------------------------------------
    Set mrsNotPayStuff = rsNotPayStuff
    With vsGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .Clear (1)
        '填充数据
        zlFullData = LoadDataToVssGrid
        .Redraw = flexRDBuffered
    End With
    
    
End Function
 
Private Sub Form_Load()
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "拒发清单"
    Call InitVsGrid
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    
    mbln按发生时间过滤 = Val(zlDatabase.GetPara("卫材医嘱按发生时间过滤", glngSys, 1723, 0))
End Sub
Private Function LoadDataToVssGrid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:将相关的数据填充到指定的网格控件中
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 11:06:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    LoadDataToVssGrid = False
    
    err = 0: On Error GoTo ErrHand:
    mlngCount恢复 = 0
    '填充数据到控件中
    mrsNotPayStuff.Filter = 0
    If mrsNotPayStuff.RecordCount <> 0 Then mrsNotPayStuff.MoveFirst
    
    With vsGrid
        If mrs拒发.RecordCount <> 0 Then mrs拒发.MoveFirst
        lngRow = .FixedRows
        Do While Not mrs拒发.EOF
            .RowData(lngRow) = Val(mrs拒发!Id)
            .TextMatrix(lngRow, .ColIndex("科室")) = NVL(mrs拒发!科室)
            .TextMatrix(lngRow, .ColIndex("开单医生")) = NVL(mrs拒发!开单医生)
            .TextMatrix(lngRow, .ColIndex("状态")) = "不处理"
            '24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料；
            .TextMatrix(lngRow, .ColIndex("单据类型")) = Decode(NVL(mrs拒发!单据), 24, "收费单", 25, "记帐单", 26, "记帐表", "不知") & IIf(mrs拒发!已收费 = 0, "(未)", "")
            .Cell(flexcpData, lngRow, .ColIndex("单据类型")) = NVL(mrs拒发!单据)
            .TextMatrix(lngRow, .ColIndex("单据号")) = NVL(mrs拒发!NO)
            .Cell(flexcpData, lngRow, .ColIndex("单据号")) = NVL(mrs拒发!费用ID)
            
            .TextMatrix(lngRow, .ColIndex("记帐员")) = NVL(mrs拒发!审核人)  '操作员姓名
            .TextMatrix(lngRow, .ColIndex("床号")) = NVL(mrs拒发!床号)
            .TextMatrix(lngRow, .ColIndex("病人姓名")) = NVL(mrs拒发!姓名)
            .TextMatrix(lngRow, .ColIndex("住院号")) = NVL(mrs拒发!住院号)
            .TextMatrix(lngRow, .ColIndex("材料名称")) = NVL(mrs拒发!品名)
            .Cell(flexcpData, lngRow, .ColIndex("材料名称")) = NVL(mrs拒发!药品ID)
            .TextMatrix(lngRow, .ColIndex("规格")) = NVL(mrs拒发!规格)
            .TextMatrix(lngRow, .ColIndex("产地")) = NVL(mrs拒发!产地)
            .TextMatrix(lngRow, .ColIndex("批号")) = NVL(mrs拒发!批号)
            .Cell(flexcpData, lngRow, .ColIndex("批号")) = NVL(mrs拒发!批次)
            
            '.TextMatrix(lngRow, .ColIndex("付")) = Format(Val(NVL(mrs拒发!付)), "###")
            .TextMatrix(lngRow, .ColIndex("数量")) = NVL(mrs拒发!数量)
            .TextMatrix(lngRow, .ColIndex("单价")) = Format(Val(NVL(mrs拒发!单价)) * mrs拒发!换算系数, mFMT.FM_零售价)
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(NVL(mrs拒发!金额)), mFMT.FM_金额)
            .TextMatrix(lngRow, .ColIndex("说明")) = NVL(mrs拒发!说明)
            .TextMatrix(lngRow, .ColIndex("记帐时间")) = NVL(mrs拒发!登记时间)
            .TextMatrix(lngRow, .ColIndex("记录性质")) = NVL(mrs拒发!记录性质)
            .TextMatrix(lngRow, .ColIndex("门诊标志")) = NVL(mrs拒发!门诊标志)
            lngRow = lngRow + 1: .Rows = .Rows + 1
            mrs拒发.MoveNext
        Loop
        
        If mrsNotPayStuff.RecordCount <> 0 Then mrsNotPayStuff.MoveFirst
        Do While Not mrsNotPayStuff.EOF
            If mrsNotPayStuff!执行状态 = 2 Then
                .RowData(lngRow) = 0
                .TextMatrix(lngRow, .ColIndex("科室")) = NVL(mrsNotPayStuff!科室)
                .TextMatrix(lngRow, .ColIndex("开单医生")) = NVL(mrsNotPayStuff!开单医生)
                .TextMatrix(lngRow, .ColIndex("状态")) = ""
                '24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料；
                .TextMatrix(lngRow, .ColIndex("单据类型")) = NVL(mrsNotPayStuff!类型)
                .TextMatrix(lngRow, .ColIndex("单据号")) = NVL(mrsNotPayStuff!NO)
                .TextMatrix(lngRow, .ColIndex("记帐员")) = NVL(mrsNotPayStuff!记帐员)
                .TextMatrix(lngRow, .ColIndex("床号")) = NVL(mrsNotPayStuff!床号)
                .TextMatrix(lngRow, .ColIndex("病人姓名")) = NVL(mrsNotPayStuff!姓名)
                .TextMatrix(lngRow, .ColIndex("住院号")) = NVL(mrsNotPayStuff!住院号)
                .TextMatrix(lngRow, .ColIndex("材料名称")) = NVL(mrsNotPayStuff!材料名称)
                .TextMatrix(lngRow, .ColIndex("规格")) = NVL(mrsNotPayStuff!规格)
                .TextMatrix(lngRow, .ColIndex("产地")) = NVL(mrsNotPayStuff!产地)
                .TextMatrix(lngRow, .ColIndex("批号")) = NVL(mrsNotPayStuff!批号)
                '.TextMatrix(lngRow, .ColIndex("付")) = Format(Val(NVL(mrsNotPayStuff!付)), "###")
                .TextMatrix(lngRow, .ColIndex("数量")) = NVL(mrsNotPayStuff!数量)
                .TextMatrix(lngRow, .ColIndex("单价")) = Format(Val(NVL(mrsNotPayStuff!单价)) * mrsNotPayStuff!换算系数, mFMT.FM_零售价)
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(NVL(mrsNotPayStuff!金额)), mFMT.FM_金额)
                .TextMatrix(lngRow, .ColIndex("说明")) = NVL(mrsNotPayStuff!说明)
                .TextMatrix(lngRow, .ColIndex("记帐时间")) = NVL(mrsNotPayStuff!记帐时间)
                .TextMatrix(lngRow, .ColIndex("记录性质")) = NVL(mrsNotPayStuff!记录性质)
                .TextMatrix(lngRow, .ColIndex("门诊标志")) = NVL(mrsNotPayStuff!门诊标志)
            
                .Cell(flexcpData, lngRow, .ColIndex("单据号")) = NVL(mrsNotPayStuff!费用ID)
                .Cell(flexcpData, lngRow, .ColIndex("单据类型")) = NVL(mrsNotPayStuff!单据)
                .Cell(flexcpData, lngRow, .ColIndex("材料名称")) = NVL(mrsNotPayStuff!材料ID)
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            mrsNotPayStuff.MoveNext
         Loop
         If .Rows > 2 Then .Rows = .Rows - 1
        If .Rows > 2 Then
            .Cell(flexcpBackColor, 1, .ColIndex("状态"), .Rows - 1, .ColIndex("状态")) = &HE7CFBA
        End If
         
    End With
    LoadDataToVssGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function RefreshData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:重新刷新拒发数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-24 10:47:45
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strWhere As String, strWhere1 As String
    Dim lngRow As Long, strFields As String
    Dim str门诊 As String
    Dim str病区发料 As String
    Dim str住院 As String
    Dim strSqlTmp As String
    
    On Error GoTo ErrHandle
    str病区发料 = zlDatabase.GetPara("病区发料方式", glngSys, mlngModule, "临床,护理,检查,检验,手术,治疗,营养")
    
    If mintUnit = 0 Then
        strFields = "X.计算单位 单位, 1 换算系数,"
    Else
        strFields = " D.包装单位 单位, D.换算系数,"
    End If
    
    strWhere = ""
    If (Trim(mArrFilter("单据号")(0)) <> "" And Trim(mArrFilter("单据号")(1)) = "") Then
        strWhere = strWhere & "            AND s.NO =[6]  "
    ElseIf (Trim(mArrFilter("单据号")(1)) <> "" And Trim(mArrFilter("单据号")(0)) = "") Then
        strWhere = strWhere & "            AND s.NO =[7]  "
    ElseIf Trim(mArrFilter("单据号")(0)) <> "" And Trim(mArrFilter("单据号")(1)) <> "" Then
        strWhere = strWhere & "            AND ( s.NO between [6] and [7] )"
    End If
    
    strWhere1 = ""
    'If Val(mArrFilter("开单科室id")) <> 0 Then strWhere1 = "  AND c.开单部门ID=[5]  "
    If Trim(mArrFilter("开单科室ID")) <> "" Then
        Select Case Val(mArrFilter("部门类型"))
        Case 0  '临床
            strWhere1 = strWhere1 & " And Instr([5], ',' || c.开单部门id || ',') > 0 And c.病人科室id=c.开单部门id"
        Case 1 '医技
            strWhere1 = strWhere1 & " And Instr([5], ',' || c.开单部门id || ',') > 0 And c.病人科室id<>c.开单部门id"
        Case Else
            '病区
            If str病区发料 = "" Then
                strWhere1 = strWhere1 & " And Instr([5], ',' || c.病人病区ID || ',') > 0 And c.病人科室id=c.开单部门id"
            Else
                strWhere1 = strWhere1 & " And Instr([5], ',' || c.病人病区ID || ',') > 0 "
                If str病区发料 <> mstrAllType Then
                    strWhere1 = strWhere1 & " And C.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                        " Where Instr([13],',' || 工作性质 || ',') > 0) "
                End If
            End If
        End Select
    End If
        
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("病人ID")) = 0, "", "  AND c.病人iD=[8]  ")
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("住院号")) = 0, "", "  AND c.标识号=[9] and c.门诊标志=2 ")
    strWhere1 = strWhere1 & IIf(Trim(mArrFilter("姓名")) = "", "", "  AND c.姓名 like [10] ")
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("门诊号")) = 0, "", "  AND C.标识号=[11] and c.门诊标志=1 ")
    strWhere1 = strWhere1 & IIf(Trim(mArrFilter("就诊卡号")) = "", "", "  AND c1.就诊卡号=[12]")
    
    
    If mbln按发生时间过滤 = False Then
        strSqlTmp = "Select S.ID, S.药品id, S.配药人, S.单据, Nvl(S.扣率, 0) As 扣率, S.NO, S.付数, S.实际数量 As 数量, S.批号, s.产地, S.批次, " & _
            " S.零售价 As 单价, S.零售金额 As 金额, S.单量, S.频次, S.用法, S.摘要 说明, S.费用id, S.对方部门id " & _
            " From 药品收发记录 S " & _
            " Where Mod(S.记录状态, 3) = 1 And Nvl(LTrim(RTrim(S.摘要)), 'Not拒发') = '拒发' And S.审核人 Is Null  " & _
            " And (S.库房id + 0 = [1] Or S.库房id Is Null)  " & strWhere & _
            " And (S.填制日期 Between [2] And [3] )  And S.单据 In (Select * From Table(Cast(f_Num2List([4]) As Zltools.t_NumList)))"
    Else
        strSqlTmp = "" & _
            "Select S.ID, S.药品id, S.配药人, S.单据, Nvl(S.扣率, 0) As 扣率, S.NO, S.付数, S.实际数量 As 数量, S.批号, s.产地, S.批次, " & _
            " S.零售价 As 单价, S.零售金额 As 金额, S.单量, S.频次, S.用法, S.摘要 说明, S.费用id, S.对方部门id " & _
            " From 药品收发记录 S, 门诊费用记录 A " & _
            " Where Mod(S.记录状态, 3) = 1 And Nvl(LTrim(RTrim(S.摘要)), 'Not拒发') = '拒发' And S.审核人 Is Null  " & _
            " And (S.库房id + 0 = [1] Or S.库房id Is Null)  " & strWhere & _
            " And (S.填制日期 Between [2] And [3] )  And S.单据 In (Select * From Table(Cast(f_Num2List([4]) As Zltools.t_NumList))) " & _
            " And S.费用id = A.Id And A.医嘱序号 Is Null "
        strSqlTmp = strSqlTmp & " Union All " & _
            "Select S.ID, S.药品id, S.配药人, S.单据, Nvl(S.扣率, 0) As 扣率, S.NO, S.付数, S.实际数量 As 数量, S.批号, s.产地, S.批次, " & _
            " S.零售价 As 单价, S.零售金额 As 金额, S.单量, S.频次, S.用法, S.摘要 说明, S.费用id, S.对方部门id " & _
            " From 药品收发记录 S, 门诊费用记录 A " & _
            " Where Mod(S.记录状态, 3) = 1 And Nvl(LTrim(RTrim(S.摘要)), 'Not拒发') = '拒发' And S.审核人 Is Null  " & _
            " And (S.库房id + 0 = [1] Or S.库房id Is Null)  " & strWhere & _
            " And (A.发生时间 Between [2] And [3] )  And S.单据 In (Select * From Table(Cast(f_Num2List([4]) As Zltools.t_NumList))) " & _
            " And S.费用id = A.Id And A.医嘱序号 Is Not Null "
    End If
    
    gstrSQL = "" & _
     "Select Distinct S.ID, S.药品id, P.名称 科室, S.配药人,C.操作员姓名 As 开单医生, C.操作员姓名 审核人, S.单据, S.扣率, S.NO, '' 床号, C.姓名,C.标识号 as 住院号, " & _
     "                 C.记录性质,C.门诊标志,C.登记时间, '[' || X.编码 || ']' || X.名称 品名, S.付数 付, S.数量, Nvl(D.在用分批, 0) 分批, X.规格, " & strFields & _
     "                Decode(S.批号, Null, '', S.批号) 批号, Nvl(S.批次, 0) 批次 , S.单价, " & _
     "                 S.金额, S.单量, S.频次, S.用法, S.说明, C.医嘱序号,S.费用ID,C.记录状态 As 已收费, Nvl(s.产地, Nvl(x.产地, '')) 产地 " & _
     " From (" & strSqlTmp & ") S, 门诊费用记录 C,病人信息 C1, 部门表 P, " & _
     "      材料特性 D, 收费项目目录 X, 收费项目别名 A  " & _
     " Where S.药品id = D.材料id And S.药品id = X.ID  And S.药品id = A.收费细目id(+) And A.性质(+) = 3  " & _
     "       And S.对方部门id = P.ID  And S.费用id = C.ID And Nvl(c.费用状态,0)<>1 and C.病人ID=C1.病人id(+)  " & vbCrLf & strWhere1
    
    '排除对未发药品的销帐记录
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From 病人费用销帐 X " & _
        " Where X.申请类别 = 0 And X.状态+0 = 0 And X.收费细目id+0 = S.药品id And X.费用id = S.费用id) "
    
    '收费处方显示方式
    If Val(mArrFilter("收费处方")) = 1 Then
        gstrSQL = gstrSQL & " And C.记录状态=1 "
    ElseIf Val(mArrFilter("收费处方")) = 2 Then
        gstrSQL = gstrSQL & " And C.记录状态=0 "
    End If
    
    If Val(mArrFilter("请求类型")) = 0 Then
        '所有
        str门诊 = Replace(gstrSQL, "c.病人病区ID", "c.开单部门id")
        str住院 = Replace(gstrSQL, "'' 床号", "c.床号")
        str住院 = Replace(str住院, "c.姓名", "nvl(R.姓名,c.姓名)")
        str住院 = Replace(str住院, "C.姓名", "nvl(R.姓名,C.姓名) 姓名")
        str住院 = Replace(str住院, "门诊费用记录 C", "住院费用记录 C,病案主页 r")
        str住院 = Replace(str住院, "And Nvl(c.费用状态,0)<>1", "and r.病人id=c.病人id and r.主页id=c.主页id " & IIf(Trim(mArrFilter("床号")) = "", "", "   AND c.床号 =[14] "))
        If Trim(mArrFilter("床号")) <> "" Then str门诊 = str门诊 & " and 1=0"
        gstrSQL = str门诊 & " Union All " & str住院
    ElseIf Val(mArrFilter("请求类型")) = 1 Then
        gstrSQL = Replace(gstrSQL, "c.病人病区ID", "c.开单部门id")
    ElseIf Val(mArrFilter("请求类型")) = 2 Then
        '住院记帐单
        gstrSQL = Replace(gstrSQL, "'' 床号", "c.床号")
        gstrSQL = Replace(gstrSQL, "c.姓名", "nvl(R.姓名,c.姓名)")
        gstrSQL = Replace(gstrSQL, "C.姓名", "nvl(R.姓名,C.姓名) 姓名")
        gstrSQL = Replace(gstrSQL, "门诊费用记录 C", "住院费用记录 C,病案主页 r")
        gstrSQL = Replace(gstrSQL, "And Nvl(c.费用状态,0)<>1", "and r.病人id=c.病人id and r.主页id=c.主页id " & IIf(Trim(mArrFilter("床号")) = "", "", "   AND c.床号 =[14] "))
    End If
     
    gstrSQL = gstrSQL & " Order By NO, 单据"
    
    Set mrs拒发 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        Val(mArrFilter("发料部门ID")), _
        CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), _
        CStr("," & mArrFilter("单据") & ","), _
        Val(mArrFilter("开单科室ID")), _
        CStr(mArrFilter("单据号")(0)), CStr(mArrFilter("单据号")(1)), _
        Val(mArrFilter("病人ID")), Val(mArrFilter("住院号")), _
        CStr(mArrFilter("姓名")), Val(mArrFilter("门诊号")), CStr(mArrFilter("就诊卡号")), "," & str病区发料 & ",", Val(mArrFilter("床号")))
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Property Get zlHaveData() As Boolean
    Dim i As Integer
    With vsGrid
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("材料名称")) <> "" Then zlHaveData = True: Exit Function
        Next
    End With
    zlHaveData = False
End Property
Public Property Get zlHaveSel恢复() As Boolean
    zlHaveSel恢复 = mlngCount恢复 > 0
End Property


Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "拒发清单"
End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        Select Case Col
        Case .ColIndex("状态")
            If zlStr.IsHavePrivs(mstrPrivs, "卫生材料恢复") = False Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Sub vsGrid_DblClick()
    Dim str状态 As String
    If zlStr.IsHavePrivs(mstrPrivs, "卫生材料恢复") = False Then Exit Sub
    
    With vsGrid
        If .Row < 1 Then
            Exit Sub
        End If
        str状态 = Trim(.TextMatrix(.Row, .ColIndex("状态")))
        If str状态 = "" Or str状态 = "拒发" Then Exit Sub
        .TextMatrix(.Row, .ColIndex("状态")) = Decode(str状态, "恢复", "不处理", "恢复")
        If .TextMatrix(.Row, .ColIndex("状态")) = "恢复" Then
            mlngCount恢复 = mlngCount恢复 + 1
        Else
            mlngCount恢复 = mlngCount恢复 - 1
        End If
    End With
End Sub



Public Sub zlSetFontSize(ByVal curFontSize As Currency)
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-05-06 17:00:44
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        .Font.Size = curFontSize
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("刘") + 120
        .RowHeightMax = TextHeight("刘") + 120
        .Refresh
    End With
End Sub




