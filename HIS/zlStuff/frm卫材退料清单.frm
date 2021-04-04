VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm卫材退料清单 
   BorderStyle     =   0  'None
   Caption         =   "退料清单"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   3900
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   7485
      _cx             =   13203
      _cy             =   6879
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
      BackColorSel    =   16711680
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm卫材退料清单.frx":0000
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
Attribute VB_Name = "frm卫材退料清单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
Private mrsBackStuff As ADODB.Recordset
Private mintUnit As Integer
Private mstrPrivs As String
Private mlngModule As Long
Private mArrFilter As Variant   '过滤条件
Private mfrmMain As Form        '父窗口
Private mbln领料人签名 As Boolean
Private mblnHave退料 As Boolean '是否勾选了退料数据的
Private Const mstrAllType As String = "临床,护理,检查,检验,手术,治疗,营养"

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private mblnFilterChange As Boolean '条件发生了改变,需要重新刷新数据
Private mbln显示整个过程 As Boolean   '显示整个过程的单据


Public Event zlRefreshDataRecordSet(ByVal rsNotStuffStuff As ADODB.Recordset)

Private mobjPlugIn As Object             '外挂接口对象

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
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
        .ColData(.ColIndex("准退数量")) = 1
        .ColData(.ColIndex("已退数量")) = 1
        
        .Cell(flexcpForeColor, 0, .ColIndex("状态")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("退料数量")) = vbBlue
        
        .Cell(flexcpFontBold, 0, .ColIndex("状态")) = True
        .Cell(flexcpFontBold, 0, .ColIndex("退料数量")) = True
    End With
End Sub


Public Function zlBackPayStuff() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:退已经发放卫生材料
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-22 14:25:18
    '-----------------------------------------------------------------------------------------------------------
    If ISValied() = False Then Exit Function
    If SaveData() = False Then Exit Function
    zlBackPayStuff = True
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
    Set mfrmMain = frmMain: mstrPrivs = strPrivs: mlngModule = lngModule
    Set mArrFilter = arrFilter
    mintUnit = intUnit
    
    '初始化值
    Call initPara
    zlRefreshData = RefreshData

End Function

 
Private Sub initPara()
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
    With vsGrid
        .Editable = flexEDKbdMouse
    End With
    
End Sub
Private Sub Form_Load()
    mblnFilterChange = True
    mbln显示整个过程 = False
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "退料清单"
    Call InitVsGrid
    Call initPara
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With vsGrid
        .Top = Me.ScaleTop
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = ScaleHeight - .Top
    End With
End Sub
Private Sub initRecStruc()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化内部记录集
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-04-24 13:16:04
    '-----------------------------------------------------------------------------------------------------------
   '已发处方记录集
    Set mrsBackStuff = New ADODB.Recordset
    With mrsBackStuff
        If .State = 1 Then .Close
        .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "类型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
        .Fields.Append "执行状态", adDouble, 1, adFldIsNullable
        .Fields.Append "记录状态", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "材料名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "分批", adDouble, 2, adFldIsNullable
        .Fields.Append "换算系数", adDouble, 3, adFldIsNullable
        .Fields.Append "付", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "已退数", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "准退数", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "退料数", adDouble, 18, adFldIsNullable
        .Fields.Append "可操作", adDouble, 2, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单价", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "金额", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单量", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "用法", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "说明", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "操作员", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "开单医生", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "发料时间", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "位置", adDouble, 18, adFldIsNullable
        .Fields.Append "医嘱id", adDouble, 18, adFldIsNullable
        .Fields.Append "领料人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "实际数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "实际价格", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 18, adFldIsNullable
        .Fields.Append "记帐员", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function RefreshData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-24 13:37:03
    '-----------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strWhere1 As String, strTemp As String, strFields As String
    Dim blnHistory As Boolean, strTable As String, strTable1 As String, rsTemp As New ADODB.Recordset
    Dim str门诊 As String
    Dim str病区发料 As String
    
    On Error GoTo ErrHandle
    str病区发料 = zlDatabase.GetPara("病区发料方式", glngSys, mlngModule, "临床,护理,检查,检验,手术,治疗,营养")
    
    If mblnFilterChange = False Then RefreshData = True: Exit Function
    
    '先检查是否存在历史数据
    blnHistory = zlDatabase.DateMoved(mArrFilter("日期范围")(0), , , Me.Caption)
    
    Select Case mintUnit
    Case 0  '散装单位
         strFields = "X.计算单位 单位,1 as 换算系数, "
    Case Else
         strFields = "D.包装单位 单位,D.换算系数,"
    End Select
    

    strWhere1 = ""
    If (Trim(mArrFilter("单据号")(0)) <> "" And Trim(mArrFilter("单据号")(1)) = "") Then
        strWhere1 = strWhere1 & "            AND A.NO =[6]  "
    ElseIf (Trim(mArrFilter("单据号")(1)) <> "" And Trim(mArrFilter("单据号")(0)) = "") Then
        strWhere1 = strWhere1 & "            AND A.NO =[7]  "
    ElseIf Trim(mArrFilter("单据号")(0)) <> "" And Trim(mArrFilter("单据号")(1)) <> "" Then
        strWhere1 = strWhere & "            AND ( A.NO between [6] and [7] )"
    End If
    
    
    'If Val(mArrFilter("开单科室id")) <> 0 Then strWhere1 = "  AND H.开单部门ID=[5]  "
    If Trim(mArrFilter("开单科室ID")) <> "" Then
        Select Case Val(mArrFilter("部门类型"))
        Case 0  '临床
            strWhere1 = strWhere1 & " And Instr([5], ',' || H.开单部门id || ',') > 0 And H.病人科室id=H.开单部门id"
        Case 1 '医技
            strWhere1 = strWhere1 & " And Instr([5], ',' || H.开单部门id || ',') > 0 And H.病人科室id<>H.开单部门id"
        Case Else
            '病区
            If str病区发料 = "" Then
                strWhere1 = strWhere1 & " And Instr([5], ',' || H.病人病区ID || ',') > 0 And H.病人科室id=H.开单部门id"
            Else
                strWhere1 = strWhere1 & " And Instr([5], ',' || H.病人病区ID || ',') > 0 "
                If str病区发料 <> mstrAllType Then
                    strWhere1 = strWhere1 & " And H.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                        " Where Instr([13],',' || 工作性质 || ',') > 0) "
                End If
            End If
        End Select
    End If
    
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("病人ID")) = 0, "", "  AND H.病人iD=[8]  ")
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("住院号")) = 0, "", "  AND H.标识号=[9] and H.门诊标志=2 ")
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("门诊号")) = 0, "", "  AND H.标识号=[11] and H.门诊标志=1 ")
    strWhere1 = strWhere1 & IIf(Trim(mArrFilter("就诊卡号")) = "", "", "  AND H1.就诊卡号=[12]  ")
 
    '获取已发料或退料的金额
    strTable = " " & _
    "   Select A.ID, A.NO, A.单据, A.序号, A.药品id, A.费用id, A.批次, A.批号, A.效期, Nvl(A.扣率, 0) 扣率, " & _
    "          Nvl(A.付数, 1) 付数, A.实际数量 实际数量, Nvl(A.付数, 1) * A.实际数量 - B.已发数量 已退数量, B.已发数量, " & _
    "          A.记录状态, A.零售价, A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门id, A.库房id, " & _
    "          A.产地, Decode(Nvl(A.领用人, ''), '', '', DECODE(mod(a.记录状态,3),2,'(退)','(领)') || A.领用人) 领料人, H.医嘱序号,H.操作员姓名, " & _
    "          H.序号 As 费用序号,H.开单人 as 开单医生,H.姓名,H.病人id,H.记录性质,H.门诊标志,H.标识号,'' 床号,1 可操作" & _
    "   From 药品收发记录 A, 门诊费用记录 H,病人信息 H1, " & _
    "(Select a.No, a.单据, a.药品id, a.序号, Sum(Nvl(a.付数, 1) * a.实际数量) 已发数量" & vbNewLine & _
    "From 药品收发记录 A," & vbNewLine & _
    "     (Select NO, 单据, 库房id,序号,药品id From 药品收发记录 Where 库房id + 0 = [1] And 审核日期 Between [2] And [3] and (记录状态 = 1 Or Mod(记录状态, 3) = 0)) B" & vbNewLine & _
    "Where a.审核人 Is Not Null And a.库房id = b.库房id And a.No = b.No and A.序号=B.序号 and A.药品id + 0 = B.药品id And a.单据 = b.单据" & vbNewLine & _
    "Group By a.No, a.单据, a.药品id, a.序号) B" & _
    "   Where A.NO = B.NO And A.单据 = B.单据 And A.药品id + 0 = B.药品id And A.序号 = B.序号  " & _
    "         And A.审核人 Is Not Null And (A.记录状态 = 1 Or Mod(A.记录状态, 3) = 0)  " & _
    "         And A.费用id = H.ID And H.病人ID=H1.病人id(+) "
    
    If mbln显示整个过程 = False Then
         strTable = strTable & " And B.已发数量 <> 0 " & vbCrLf & strWhere1
        
        If blnHistory Then
            strTable = AnalyseHistorySQL(strTable, "1 可操作", "-99 可操作")
        End If
       
        gstrSQL = " " & _
        "   Select /*+ cardinality(J,10)*/ Distinct S.ID, S.单据, S.药品id, S.NO, S.序号, S.扣率, P.名称 科室,S.记录性质,S.门诊标志, S.标识号, S.病人id, S.床号,S.操作员姓名, " & _
        "                   S.姓名, '[' || X.编码 || ']' || X.名称 品名, Nvl(D.在用分批, 0) 分批, X.规格," & strFields & _
        "                   S.付数 付, S.实际数量 数量, S.已退数量, S.已发数量 准退数, " & _
        "                   Decode(S.批号, Null, '', S.批号)  批号, " & _
        "                   Nvl(S.批次, 0) 批次, S.效期, S.零售价 单价, S.零售金额 金额, S.单量, S.频次, S.用法, S.摘要 说明, " & _
        "                   S.审核人, To_Char(S.审核日期, 'YYYY-MM-DD HH24:MI:SS') 发料时间, 1 可操作, S.医嘱序号, I.计算单位, " & _
        "                   Nvl(S.产地, Nvl(X.产地, '')) 产地, Nvl(M.审查结果, -1) 审查结果, Nvl(S.医嘱序号, -1) 医嘱id, S.领料人, " & _
        "                   '' 库房货位, M.相关id, S.费用序号, Z.名称 As 其它名,S.记录状态,S.开单医生 " & _
        "   From (" & vbCrLf & strTable & vbCrLf & ") S, 部门表 P,Table(Cast(f_Num2List([4]) As zlTools.t_NumList)) J, " & _
        "        材料特性 D, 收费项目目录 X, 收费项目别名 A, 诊疗项目目录 I, 病人医嘱记录 M, 诊疗项目别名 Z " & _
        "   Where S.药品id = D.材料id And S.对方部门id + 0 = P.ID And D.材料id = X.ID And D.诊疗id = I.ID And S.医嘱序号 = M.ID(+) And " & _
        "         D.诊疗id = Z.诊疗项目id(+) And Z.性质(+) = 2 And D.材料id = A.收费细目id(+) And A.性质(+) = 3 And S.单据 =J.Column_Value  And " & _
        "         (S.记录状态 = 1 Or Mod(S. 记录状态, 3) = 0) And S.审核人 Is Not Null And S.库房id + 0 = [1]  And " & _
        "         S.实际数量 * S.付数 > S.已退数量 "
    Else
        '清单显示每笔操作过程
        strTable = strTable & strWhere1
        If blnHistory Then
            strTable = AnalyseHistorySQL(strTable, "1 可操作", "-99 可操作")
        End If
        
        strTable1 = " Union All " & _
        "     Select A.ID, A.NO, A.单据, A.序号, A.药品id, A.费用id, A.批次, A.批号, A.效期, Nvl(A.扣率, 0), Nvl(A.付数, 1) 付数, " & _
        "            A.实际数量, 0 已退数, 0 已发数量, A.记录状态, A.零售价, A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, " & _
        "            A.审核日期, A.对方部门id, A.库房id, " & _
        "            A.产地, " & _
        "            Decode(Nvl(A.领用人, ''), '', '',Decode(A.记录状态, 2,'(退)', '(领)' )|| A.领用人) 领料人,H.医嘱序号,H.操作员姓名, " & _
        "          H.序号 As 费用序号,H.开单人 as 开单医生,H.姓名,H.病人id,H.记录性质,H.门诊标志,H.标识号,'' 床号, Decode(A.记录状态, 1, 1,Mod(A.记录状态, 3) + 1) 可操作 " & _
        "     From 药品收发记录 A, 门诊费用记录 H ,病人信息 H1" & _
        "     Where A.费用id=H.id And H.病人id=H1.病人ID(+) and A.审核人 Is Not Null And Not (A.记录状态 = 1 Or Mod(A.记录状态, 3) = 0) And A.库房id + 0 = [1] And " & _
        "           A.审核日期 Between [2] And [3] " & strWhere1
        If blnHistory Then
            '历史数据，不能操作
            strTable1 = AnalyseHistorySQL(strTable1, "Decode(A.记录状态, 1, 1,Mod(A.记录状态, 3) + 1) 可操作", "-99 可操作")
        End If
        
        strTable = strTable & vbCrLf & strTable1
        gstrSQL = " " & _
        "     Select /*+ cardinality(J,10)*/ Distinct S.ID, S.单据, S.药品id, S.NO, S.序号, S.扣率, P.名称 科室, S.记录性质,S.门诊标志, S.标识号, S.病人id, S.床号,S.操作员姓名, " & _
        "                     S.姓名, '[' || X.编码 || ']' || X.名称 品名, Nvl(D.在用分批, 0) 分批, X.规格, " & strFields & _
        "                     S.付数 付, S.实际数量 数量, S.已退数量, S.已发数量 准退数, " & _
        "                     Decode(S.批号, Null, '', S.批号)  批号, " & _
        "                     Nvl(S.批次, 0) 批次, S.效期, S.零售价 单价, S.零售金额 金额, S.单量, S.频次, S.用法, S.摘要 说明, " & _
        "                     To_Char(S.审核日期, 'YYYY-MM-DD HH24:MI:SS') 发料时间, S.审核人, S.审核日期, 可操作, S.医嘱序号, " & _
        "                     I.计算单位, Nvl(S.产地, Nvl(X.产地, '')) 产地, Nvl(M.审查结果, -1) 审查结果, " & _
        "                     Nvl(S.医嘱序号, -1) 医嘱id, S.领料人, '' 库房货位, Z.名称 As 其它名,S.记录状态,s.开单医生 " & _
        "     From (" & strTable & ") S, 部门表 P, 材料特性 D, 收费项目目录 X,Table(Cast(f_Num2List([4]) As zlTools.t_NumList)) J, " & _
        "          收费项目别名 A, 诊疗项目目录 I, 病人医嘱记录 M, 诊疗项目别名 Z " & _
        "     Where S.药品id = D.材料id And D.材料id = X.ID And S.对方部门id + 0 = P.ID And D.诊疗id = I.ID And " & _
        "           S.医嘱序号 = M.ID(+) And D.诊疗id = Z.诊疗项目id(+) And Z.性质(+) = 2 And D.材料id = A.收费细目id(+) And " & _
        "           A.性质(+) = 3 And  S.单据 =J.Column_Value And S.审核人 Is Not Null "
    End If
    
    If Val(mArrFilter("请求类型")) = 0 Then
        '所有
        str门诊 = Replace(gstrSQL, "H.病人病区ID", "H.开单部门ID")
        str门诊 = str门诊 & IIf(Trim(mArrFilter("姓名")) = "", "", "  AND S.姓名 like [10] ")
        
        gstrSQL = Replace(gstrSQL, "'' 床号", "H.床号")
        gstrSQL = Replace(gstrSQL, "S.姓名", "R.姓名")
        gstrSQL = Replace(gstrSQL, "H.姓名", "H.主页id")
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        gstrSQL = Replace(gstrSQL, "病人医嘱记录 M", "病人医嘱记录 M,病案主页 r")
        gstrSQL = gstrSQL & " and r.病人id=S.病人id and r.主页id=S.主页id " & IIf(Trim(mArrFilter("床号")) = "", "", "   AND S.床号 =[14] ")
        gstrSQL = gstrSQL & IIf(Trim(mArrFilter("姓名")) = "", "", "  AND R.姓名 like [10] ")
        If Trim(mArrFilter("床号")) <> "" Then str门诊 = str门诊 & " and 1=0"
        gstrSQL = str门诊 & " Union All " & gstrSQL
    ElseIf Val(mArrFilter("请求类型")) = 2 Then
        '住院记帐单
        gstrSQL = Replace(gstrSQL, "'' 床号", "H.床号")
        gstrSQL = Replace(gstrSQL, "H.姓名", "H.主页id")
        gstrSQL = Replace(gstrSQL, "S.姓名", "R.姓名")
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        gstrSQL = Replace(gstrSQL, "病人医嘱记录 M", "病人医嘱记录 M,病案主页 r")
        gstrSQL = gstrSQL & " and r.病人id=S.病人id and r.主页id=S.主页id " & IIf(Trim(mArrFilter("床号")) = "", "", "   AND S.床号 =[14] ")
        gstrSQL = gstrSQL & IIf(Trim(mArrFilter("姓名")) = "", "", "  AND R.姓名 like [10] ")
    End If
    
    If mbln显示整个过程 = False Then
        gstrSQL = gstrSQL & " Order By NO, 单据, 费用序号"
    Else
        gstrSQL = gstrSQL & " Order By NO, 单据, 审核日期"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        Val(mArrFilter("发料部门ID")), _
        CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), _
        CStr(mArrFilter("单据")), _
        "," & Trim(mArrFilter("开单科室ID")) & ",", _
        CStr(mArrFilter("单据号")(0)), CStr(mArrFilter("单据号")(1)), _
        Val(mArrFilter("病人ID")), Val(mArrFilter("住院号")), _
        CStr(mArrFilter("姓名")), Val(mArrFilter("门诊号")), CStr(mArrFilter("就诊卡号")), "," & str病区发料 & ",", Val(mArrFilter("床号")))
    Call WhiteDataToRecord(rsTemp)
    With vsGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .Clear (1)
        '填充数据
        Call FullDataToVsGrid
        .Redraw = flexRDBuffered
    End With
    mblnHave退料 = False
    mblnFilterChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FullDataToVsGrid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:将相关的数据填充到指定的网格控件中
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 11:06:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    FullDataToVsGrid = False
    
    err = 0: On Error GoTo ErrHand:

    '填充数据到控件中
    If mrsBackStuff.RecordCount <> 0 Then mrsBackStuff.MoveFirst
    With vsGrid
        .Clear (1)
        If mrsBackStuff.EOF Then '
            .Rows = 2
            FullDataToVsGrid = True
            Exit Function
        End If
        
        .Rows = mrsBackStuff.RecordCount + .FixedRows
        lngRow = .FixedRows
        Do While Not mrsBackStuff.EOF
            .RowData(lngRow) = Val(mrsBackStuff!Id)
            .TextMatrix(lngRow, .ColIndex("科室")) = NVL(mrsBackStuff!科室)
            .TextMatrix(lngRow, .ColIndex("开单医生")) = NVL(mrsBackStuff!开单医生)
            .TextMatrix(lngRow, .ColIndex("状态")) = IIf(Val(NVL(mrsBackStuff!执行状态)) = 1, "不处理", "退料")
            .TextMatrix(lngRow, .ColIndex("单据类型")) = NVL(mrsBackStuff!类型)
            .TextMatrix(lngRow, .ColIndex("单据号")) = NVL(mrsBackStuff!NO)
            .TextMatrix(lngRow, .ColIndex("记帐员")) = NVL(mrsBackStuff!记帐员)
            .TextMatrix(lngRow, .ColIndex("床号")) = NVL(mrsBackStuff!床号)
            .TextMatrix(lngRow, .ColIndex("病人姓名")) = NVL(mrsBackStuff!姓名)
            .TextMatrix(lngRow, .ColIndex("住院号")) = NVL(mrsBackStuff!住院号)
            .TextMatrix(lngRow, .ColIndex("材料名称")) = NVL(mrsBackStuff!材料名称)
            .TextMatrix(lngRow, .ColIndex("规格")) = NVL(mrsBackStuff!规格)
            .TextMatrix(lngRow, .ColIndex("产地")) = NVL(mrsBackStuff!产地)
            .TextMatrix(lngRow, .ColIndex("批号")) = NVL(mrsBackStuff!批号) & IIf(Val(NVL(mrsBackStuff!批次)) = 0, "", "(" & NVL(mrsBackStuff!批次) & ")")
            .TextMatrix(lngRow, .ColIndex("数量")) = NVL(mrsBackStuff!数量)
            .TextMatrix(lngRow, .ColIndex("已退数量")) = NVL(mrsBackStuff!已退数)
            .TextMatrix(lngRow, .ColIndex("准退数量")) = NVL(mrsBackStuff!准退数)
            .TextMatrix(lngRow, .ColIndex("退料数量")) = IIf(Val(NVL(mrsBackStuff!执行状态)) = 1, Format("0", mFMT.FM_数量), NVL(mrsBackStuff!准退数))
            .TextMatrix(lngRow, .ColIndex("单价")) = Format(Val(NVL(mrsBackStuff!单价)) * mrsBackStuff!换算系数, mFMT.FM_零售价)
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(NVL(mrsBackStuff!金额)), mFMT.FM_金额)
            .TextMatrix(lngRow, .ColIndex("操作员")) = NVL(mrsBackStuff!操作员)
            .TextMatrix(lngRow, .ColIndex("发料时间")) = NVL(mrsBackStuff!发料时间)
            .TextMatrix(lngRow, .ColIndex("领/退料人")) = NVL(mrsBackStuff!领料人)
            
            .Cell(flexcpData, lngRow, .ColIndex("单据号")) = Val(NVL(mrsBackStuff!位置))
            .Cell(flexcpData, lngRow, .ColIndex("病人姓名")) = Val(NVL(mrsBackStuff!可操作))
            .Cell(flexcpData, lngRow, .ColIndex("状态")) = Val(NVL(mrsBackStuff!记录状态))
            SetGRDCOLOR vsGrid, lngRow, IIf(NVL(mrsBackStuff!可操作) = 1, 1, NVL(mrsBackStuff!记录状态, 0))
            lngRow = lngRow + 1
            mrsBackStuff.MoveNext
         Loop

         .Cell(flexcpFontBold, 1, .ColIndex("状态"), .Rows - 1, .ColIndex("状态")) = True
         .Cell(flexcpFontBold, 1, .ColIndex("退料数量"), .Rows - 1, .ColIndex("退料数量")) = True
    End With
    FullDataToVsGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function WhiteDataToRecord(ByVal rsSource As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:将相关的数据写入内部记录集(未发料部分)
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 10:03:41
    '-----------------------------------------------------------------------------------------------------------
    Dim ArrayPhysic
    Dim IntArray As Integer
     
    err = 0: WhiteDataToRecord = False
    
    Call initRecStruc
    
    With rsSource
        Do While Not .EOF
            mrsBackStuff.AddNew
            mrsBackStuff!Id = !Id
            mrsBackStuff!材料ID = !药品ID
            mrsBackStuff!位置 = .AbsolutePosition
            mrsBackStuff!科室 = !科室
            mrsBackStuff!类型 = Decode(NVL(!单据), 24, "收费单", 25, "记帐单", 26, "记帐表", "不知")
            mrsBackStuff!执行状态 = 1                        '缺省为不处理
            mrsBackStuff!NO = !NO
            mrsBackStuff!单据 = !单据
            mrsBackStuff!开单医生 = NVL(!开单医生)
            mrsBackStuff!序号 = !序号
            mrsBackStuff!病人ID = Val(NVL(!病人ID))
            mrsBackStuff!床号 = !床号
            mrsBackStuff!姓名 = IIf(IsNull(!姓名), "", !姓名)
            mrsBackStuff!住院号 = IIf(Val(NVL(!门诊标志)) = 2, NVL(!标识号), "")
            mrsBackStuff!材料名称 = !品名
            mrsBackStuff!规格 = IIf(IsNull(!规格), "", !规格)
            mrsBackStuff!产地 = IIf(IsNull(!产地), "", !产地)
            mrsBackStuff!分批 = IIf(IsNull(!分批), 0, !分批)
            mrsBackStuff!批次 = IIf(IsNull(!批次), 0, !批次)
            mrsBackStuff!批号 = IIf(IsNull(!批号), "", !批号)
            mrsBackStuff!效期 = IIf(IsNull(!效期), "", !效期)
            mrsBackStuff!换算系数 = !换算系数
            mrsBackStuff!付 = IIf(IsNull(!付), 1, !付)
            mrsBackStuff!数量 = Format(Val(NVL(!数量)) / !换算系数, mFMT.FM_数量) & !单位
            mrsBackStuff!已退数 = Format(Val(NVL(!已退数量)) / !换算系数, mFMT.FM_数量)
            mrsBackStuff!准退数 = Format(Val(NVL(!准退数)) / !换算系数, mFMT.FM_数量)
            mrsBackStuff!退料数 = Format(Val(NVL(!准退数)) / !换算系数, mFMT.FM_数量)
            mrsBackStuff!单位 = !单位
            mrsBackStuff!单价 = Format(!单价 * !换算系数, mFMT.FM_数量)
            mrsBackStuff!金额 = !金额
            mrsBackStuff!单量 = IIf(IsNull(!单量), "", zlStr.FormatEx(!单量, 5) & NVL(!计算单位))
           ' mrsBackStuff!单量单位 = NVL(!计算单位)
            mrsBackStuff!频次 = NVL(!频次)
            mrsBackStuff!用法 = NVL(!用法)
            mrsBackStuff!说明 = NVL(!说明)
            mrsBackStuff!操作员 = NVL(!审核人)
            mrsBackStuff!发料时间 = NVL(!发料时间)
            mrsBackStuff!可操作 = IIf(Val(NVL(!准退数)) = 0, 0, IIf(IsNull(!可操作), 0, !可操作))
            mrsBackStuff!医嘱id = !医嘱id
            mrsBackStuff!领料人 = !领料人
            mrsBackStuff!实际数量 = !准退数
            mrsBackStuff!实际价格 = !单价
            mrsBackStuff!记录状态 = Val(NVL(!记录状态))
            mrsBackStuff!记录性质 = Val(NVL(!记录性质))
            mrsBackStuff!门诊标志 = Val(NVL(!门诊标志))
            mrsBackStuff!记帐员 = NVL(!操作员姓名)
            .MoveNext
        Loop
    End With
    If err <> 0 Then
        MsgBox "产生内部记录集时，发生不可预知的错误！", vbInformation, gstrSysName
        Call initRecStruc
        Exit Function
    End If
    RaiseEvent zlRefreshDataRecordSet(mrsBackStuff)
    WhiteDataToRecord = True
End Function


Private Function AnalyseHistorySQL(ByVal strSQL As String, Optional str原串 As String = "", Optional str现串 As String = "") As String
    '产生历史数据的SQL语句
    Dim strTemp As String
    strTemp = Replace(strSQL, "药品收发记录", "H药品收发记录")
    strTemp = Replace(strTemp, "门诊费用记录", "H门诊费用记录")
    strTemp = Replace(strTemp, "住院费用记录", "H住院费用记录")
    If str原串 <> "" Then
        strTemp = Replace(strTemp, str原串, str现串)
    End If
    strTemp = strSQL & " Union ALL " & strTemp
    AnalyseHistorySQL = strTemp
End Function
Private Function ISValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查发料是否合法
    '入参:
    '出参:
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 14:25:36
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Integer, lng材料ID As Long, rsCheck As ADODB.Recordset, dbl现价 As Double
    Dim str序号 As String, blnHaveData As Boolean
    
    On Error GoTo ErrHandle
    ISValied = False
    
    '先初始检查
    Set rsCheck = CheckBillStruct
    lng材料ID = 0
    '检查执行库房
    With mrsBackStuff
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If .EOF Then Exit Function
        .Sort = "材料ID Asc"
        blnHaveData = False
        Do While Not .EOF
            If Val(NVL(!执行状态)) = 3 And Val(NVL(!退料数)) <> 0 Then
                '主要是提供检查速度，先处理内部数据集
                rsCheck.Filter = "单据标识='" & NVL(!NO) & "|" & NVL(!单据) & "'"
                If rsCheck.RecordCount <> 0 Then
                    rsCheck.Find "病人ID=" & Val(NVL(!病人ID))
                    If rsCheck.EOF Then rsCheck.AddNew
                Else
                    rsCheck.AddNew
                End If
                rsCheck!单据标识 = NVL(!NO) & "|" & NVL(!单据)
                rsCheck!病人ID = Val(NVL(!病人ID))
                rsCheck!记录性质 = Val(NVL(!记录性质))
                rsCheck!门诊标志 = Val(NVL(!门诊标志))
                str序号 = NVL(rsCheck!序号)
                If InStr(1, "," & str序号 & ",", "," & Val(NVL(!序号)) & ",") = 0 Then
                    If str序号 = "" Then
                        str序号 = Val(NVL(!序号))
                    Else
                        str序号 = str序号 & "," & Val(NVL(!序号))
                    End If
                    rsCheck!序号 = str序号
                End If
                rsCheck.Update
                rsCheck.Filter = 0
                 '如果原来不分批而现在分批
                If Val(NVL(!批次)) = 0 And Val(NVL(!分批)) = 1 Then
                    ShowMsgBox "卫材：" & NVL(!材料名称) & "原来没有分批,而现在是分批的,操作中止!"
                    Exit Function
                End If
                
                '需要检查原价与现价是否一至
                If lng材料ID <> Val(NVL(!材料ID)) Then
                
                    gstrSQL = "" & _
                        "   Select  b.现价, Nvl(C.是否变价, 0) 是否变价 " & _
                        "   From  收费价目 b, 收费项目目录 C " & _
                        "   where   b.收费细目ID=C.id and  (SYSDATE BETWEEN b.执行日期 AND b.终止日期 Or  SYSDATE >= b.执行日期 AND b.终止日期 IS Null)" & _
                        GetPriceClassString("B") & " And C.id=[1]"
                        
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "[取原始价格和最新价格]", Val(NVL(!材料ID)))
                End If
                If rsTemp.EOF Then
                    dbl现价 = Val(NVL(!实际价格))
                ElseIf Val(NVL(rsTemp!是否变价)) = 1 Then
                    dbl现价 = Val(NVL(!实际价格))
                Else
                    dbl现价 = Val(NVL(rsTemp!现价))
                End If
                If dbl现价 <> Val(NVL(!实际价格)) Then
                    If MsgBox("材料[" & !材料名称 & "(" & !规格 & ")]" & "原价为" & Val(NVL(!实际价格)) & ",现价为" & dbl现价 & "。" & vbCrLf & Space(4) & "退药将产生调价退料明细记录，是否继续退料? ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then
                        !执行状态 = 0
                        .Update
                    End If
                End If
                blnHaveData = True
            End If
            .MoveNext
        Loop
    End With
    If blnHaveData = False Then
        ShowMsgBox "未选择需要退料的卫生材料，请查检!"
        Exit Function
    End If
    Dim strNo As String, lng单据 As Long, lng病人id As Long
    '检查单据,主要是检查处方是否已经结帐,病人是否已经出院，差对权限进行相关的检查
    With rsCheck
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !单据标识 & "|"
            lng单据 = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            lng病人id = !病人ID
            str序号 = NVL(!序号)
            
            '检查结帐处方是否能发料
            If Check结帐处方(mstrPrivs, lng单据, strNo, str序号, Val(!记录性质), Val(!门诊标志)) = False Then Exit Function
            If Check出院病人(mstrPrivs, lng单据, strNo, Val(!记录性质), Val(!门诊标志), lng病人id) = False Then Exit Function
            .MoveNext
        Loop
    End With
    ISValied = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 

Private Function CheckBillStruct() As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化检查对象集
    '入参:
    '出参:
    '返回:成功,返回空记录集结构
    '编制:刘兴洪
    '日期:2008-04-23 14:41:41
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    
    With rsTemp
        If .State = 1 Then .Close
        .Fields.Append "单据标识", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set CheckBillStruct = rsTemp
End Function
Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:对指定的发料项目进行退料处理
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 11:48:06
    '-----------------------------------------------------------------------------------------------------------
    Dim strDate As String, str退料人 As String, lng病人id As Long, strID批次 As String, dbl退料数 As Double
    Dim int自动销帐 As Integer
    Dim cllPro As Collection
    Dim strReturnInfo As String
    Dim strReserve As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln备货卫材 As Boolean
    Dim int自动销帐_原始值 As Integer
    
    int自动销帐_原始值 = IIf(Val(zlDatabase.GetPara("自动销帐", glngSys, mlngModule)) = 1, 1, 0)
     
    SaveData = False
    err = 0: On Error GoTo ErrHand:
    strDate = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    Set cllPro = New Collection
    
    With mrsBackStuff
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If .EOF Then Exit Function
        
        If MsgBox("你现在确定要进行退料操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        '领药人签名
        str退料人 = ""
        If mbln领料人签名 Then
            str退料人 = zlDatabase.UserIdentify(Me, "领料人签名", glngSys, mlngModule, "退料")
            If str退料人 = "" Then
                Exit Function
            End If
        End If
        
        '必须按病人ID，材料ID排序
        .Sort = "材料ID Asc"
        Do While Not .EOF
            If !执行状态 = 3 And Val(NVL(!退料数)) <> 0 Then
                dbl退料数 = Val(!退料数) * !换算系数
                If Val(!准退数) = Val(!退料数) Then
                    dbl退料数 = Val(!实际数量)
                End If
                If dbl退料数 <> 0 Then
                    int自动销帐 = int自动销帐_原始值
                    
                    If int自动销帐 <> 1 Then
                        '判断是否备货卫材
                        gstrSQL = " Select 1 From 药品收发记录 Where 单据 = 21 And 审核日期 Is Not Null And 费用id = (select 费用id from 药品收发记录 where id=[1]) And Rownum < 2 "
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否备货卫材", NVL(!Id))
                        bln备货卫材 = Not rsTemp.EOF
                        
                        '如果是高值卫材也进行自动销帐
                        If bln备货卫材 Then int自动销帐 = 1
                    End If
                    
                    'Zl_材料收发记录_部门退料
                    gstrSQL = "Zl_材料收发记录_部门退料("
                    '    收发id_In   In 药品收发记录.ID%Type,
                    gstrSQL = gstrSQL & "" & NVL(!Id) & ","
                    '    审核人_In   In 药品收发记录.审核人%Type,
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    '    审核日期_In In 药品收发记录.审核日期%Type,
                    gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:mi:ss'),"
                    '    批号_In     In 药品库存.上次批号%Type := Null,
                        gstrSQL = gstrSQL & "'" & NVL(!批号) & "',"
                    '    效期_In     In 药品库存.效期%Type := Null,
                    gstrSQL = gstrSQL & "" & IIf(IsNull(!效期), "NULL", IIf(NVL(!效期) = "", "NULL", "To_Date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & ","
                    '    产地_In     In 药品库存.上次产地%Type := Null,
                    gstrSQL = gstrSQL & "'" & NVL(!产地) & "',"
                    '    退料数量_In In 药品收发记录.实际数量%Type := Null,
                    gstrSQL = gstrSQL & "" & dbl退料数 & ","
                    '    自动销帐_In Integer := 0,
                    gstrSQL = gstrSQL & "" & int自动销帐 & ","
                    '    退料人_In   In 药品收发记录.领用人%Type := Null
                    gstrSQL = gstrSQL & "'" & str退料人 & "')"
                    AddArray cllPro, gstrSQL
                    
                    strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & NVL(!Id) & "," & dbl退料数
                End If
                
            End If
            .MoveNext
        Loop
    End With
        
    On Error GoTo ErrExcute:
    Call ExecuteProcedureArrAy(cllPro, Me.Caption)
    SaveData = True
    
    err = 0: On Error GoTo ErrHand:
    If zlStr.IsHavePrivs(mstrPrivs, "退料通知单") Then
          If MsgBox("你需要打印退料清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
              Call zlPrintBill(False, strDate)
          End If
    End If
    
    '调用退药后的外挂接口
    If Not mobjPlugIn Is Nothing Then
        mobjPlugIn.DrugReturnByID Val(mArrFilter("发料部门id")), strReturnInfo, CDate(strDate), strReserve
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
ErrExcute:
      gcnOracle.RollbackTrans
      If ErrCenter = 1 Then Resume
      Call SaveErrLog
End Function

Public Sub zlPrintBill(ByVal bln发料单 As Boolean, Optional str发料日期 As String = "", Optional int格式 As Integer = 1, Optional strPrivs As String = "", Optional blnPrintAsk As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '功能:找印单据
    '入参:bln发料单-是否打印已经发料的发料单
    '     strDate-发料日期
    '     int格式-单据打印格式
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-06 10:39:42
    '-----------------------------------------------------------------------------------------------------------
    If mstrPrivs = "" Then mstrPrivs = strPrivs
    
    err = 0: On Error GoTo ErrHand:
    If str发料日期 = "" And blnPrintAsk = False Then
        With vsGrid
            str发料日期 = Trim(.TextMatrix(.Row, .ColIndex("发料时间")))
        End With
    End If
    If bln发料单 Then
        Call PrintPayBill(str发料日期, int格式)
        Exit Sub
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "退料通知单") Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_2", Me, "退料时间=" & str发料日期, "单位=" & mintUnit + 1, 2)
    Else
        ShowMsgBox "你不具备打印退料通知单的相关权限,请与系统管理员联系!"
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub PrintPayBill(Optional str发料日期 As String = "", Optional int格式 As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '功能:单据或清册打印
    '入参:
    '     intStyle:0-按发料方式打印,1-单据打印,2-退料单据
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-05 10:36:44
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim bln已发料清单 As Boolean
    Dim intMsg As Integer   '0-提示打印,1-自动打印,2-不打印
    
    intMsg = Val(zlDatabase.GetPara("发料打印提醒方式", glngSys, mlngModule, "0"))
 
    bln已发料清单 = zlStr.IsHavePrivs(mstrPrivs, "打印已发料清单")
    If bln已发料清单 = False Then
        ShowMsgBox "你不具备打印退料通知单的相关权限,请与系统管理员联系!"
        Exit Sub
    End If
    If intMsg = 0 Then
        '提示打印
        If MsgBox("你需要打印相关单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    ElseIf intMsg = 1 Then
        '自动打印
    Else
        Exit Sub
    End If
    '部门发料
    If str发料日期 = "" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, _
           "库房=" & Val(mArrFilter("发料部门ID")), _
           "发料方式=部门发料|3", _
           "部门性质=" & Val(mArrFilter("部门类型")), _
           "接收科室=" & 获取接收部门条件(str发料日期), _
           "单位=" & IIf(mintUnit = 0, 0, 1), _
           "ReportFormat=" & int格式, "PrintEmpty=0", 2)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, _
           "库房=" & Val(mArrFilter("发料部门ID")), _
           "发料方式=部门发料|3", _
           "部门性质=" & Val(mArrFilter("部门类型")), _
           "接收科室=" & 获取接收部门条件(str发料日期), _
           "单位=" & IIf(mintUnit = 0, 0, 1), _
           "发料号=" & str发料日期, _
           "ReportFormat=" & int格式, "PrintEmpty=0", 2)
    End If
End Sub
 
Public Function CheckPrice(ByVal lngBillId As Long, ByRef strMsg As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '判断售价是否是当前最新售价
    
    On Error GoTo ErrHandle
    '取原始价格和现价
    gstrSQL = "select nvl(a.零售价,0) 原价,b.现价, Nvl(C.是否变价, 0) 是否变价 " & _
        " from 药品收发记录 a,收费价目 b, 收费项目目录 C " & _
        " where a.药品id=b.收费细目id And A.药品id = C.ID  And (SYSDATE BETWEEN b.执行日期 AND b.终止日期 Or  SYSDATE >= b.执行日期 AND b.终止日期 IS Null)" & _
        GetPriceClassString("B") & " And a.id=[1]"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "[取原始价格和最新价格]", lngBillId)
    
    If rsTemp.RecordCount = 0 Then
        CheckPrice = True
        Exit Function
    End If
    
    '时价药品不处理
    If rsTemp!是否变价 = 1 Then
        CheckPrice = True
        Exit Function
    End If
    
    '比较价格
    If rsTemp!原价 <> rsTemp!现价 Then
        strMsg = "原价为" & rsTemp!原价 & ",现价为" & rsTemp!现价 & "。" & vbCrLf & Space(4) & "退药将产生调价退药明细记录，是否继续退药? "
        CheckPrice = False
        Exit Function
    End If
    
    CheckPrice = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Property Get zlHaveData() As Boolean
    If mrsBackStuff Is Nothing Then zlHaveData = False: Exit Sub
    zlHaveData = mrsBackStuff.RecordCount <> 0
End Property
Public Property Get zlHaveSel退料() As Boolean
    zlHaveSel退料 = mblnHave退料
End Property
Public Property Get zl显示整过程单据() As Boolean
    zl显示整过程单据 = mbln显示整个过程
End Property
Public Property Let zl显示整过程单据(ByVal vNewValue As Boolean)
    If vNewValue <> mbln显示整个过程 Then
        mbln显示整个过程 = vNewValue
        mblnFilterChange = True
         With vsGrid
            .Redraw = flexRDNone
            .Rows = .FixedRows + 1
            .Clear (1)
            '填充数据
            Call RefreshData
            Call FullDataToVsGrid
            .Redraw = flexRDBuffered
        End With
    End If
End Property
Public Property Get zlArrFilter() As Variant
    Set zlArrFilter = mArrFilter
End Property

Public Property Set zlArrFilter(ByVal vNewValue As Variant)
    Set mArrFilter = vNewValue
    mblnFilterChange = True
End Property

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "退料清单"
End Sub

Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGrid
        Select Case Col
            Case .ColIndex("状态")
            
                '改变执行状态
                Call SetExecuteStaut(Row, 0)
            Case .ColIndex("退料数量")
                Call SetExecuteStaut(Row, 1)
            Case Else
        End Select
    End With
End Sub

 Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        Select Case Col
        Case .ColIndex("状态")
            If zlStr.IsHavePrivs(mstrPrivs, "卫生材料退料") = False Then Cancel = True: Exit Sub
            '历史数据是不能更改的
            If Trim(.TextMatrix(Row, .ColIndex("病人姓名"))) = -99 Then Cancel = True
            '冲销部分也不能冲销
            If Val(.Cell(flexcpData, Row, .ColIndex("病人姓名"))) <> 1 Then Cancel = True
        Case .ColIndex("退料数量")
            If Val(.TextMatrix(Row, .ColIndex("准退数量"))) = 0 Then Cancel = True
        Case Else
            Cancel = True
        End Select
    End With
End Sub




Private Sub vsGrid_DblClick()
    Dim str状态 As String
    If zlStr.IsHavePrivs(mstrPrivs, "卫生材料退料") = False Then Exit Sub
    
    With vsGrid
        If .Row < 1 Then
            Exit Sub
        End If
        
        If .Col = .ColIndex("退料数量") Then Exit Sub
        
        If Val(.Cell(flexcpData, .Row, .ColIndex("病人姓名"))) <> 1 Then Exit Sub

        str状态 = Trim(.TextMatrix(.Row, .ColIndex("状态")))
        .TextMatrix(.Row, .ColIndex("状态")) = Decode(str状态, "退料", "不处理", "退料")
        Call SetExecuteStaut(.Row, 0)
    End With
End Sub

Private Sub SetExecuteStaut(ByVal lngRow As Long, ByVal intType As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置执行状态
    '入参:lngRow-指定的行;intType-0:更新状态;1-更新退料数量
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-04-23 11:31:04
    '-----------------------------------------------------------------------------------------------------------
    Dim str状态 As String, int状态 As Integer, lng位置 As Long
    With vsGrid
        str状态 = Trim(.TextMatrix(lngRow, .ColIndex("状态")))
        int状态 = Decode(str状态, "退料", 3, 0)
        lng位置 = Val(.Cell(flexcpData, lngRow, .ColIndex("单据号")))
    End With
    With mrsBackStuff
         .Filter = 0
        .MoveFirst
        .Find "位置=" & lng位置
        If .EOF = False Then
            If intType = 0 Then
                !执行状态 = int状态:
                If int状态 = 3 Then
                    !退料数 = !准退数
                Else
                    !退料数 = 0
                End If
            Else
                '退料数小于等于0或者大于准退数量，则不能退料，状态标志为“不处理”
                If Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("退料数量"))) <= 0 Or Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("退料数量"))) > vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("准退数量")) Then
                    vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("退料数量")) = "0"
                    int状态 = 0
                Else
                    int状态 = 3
                End If
                
                !执行状态 = int状态
                If int状态 = 3 Then
                    !退料数 = Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("退料数量")))
                Else
                    !退料数 = 0
                End If
            End If
            
            .Update
            
            vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("退料数量")) = Format(Val(!退料数), mFMT.FM_数量)
        End If
        '可能库存检查后，需要处理当前的状态
        .MoveFirst
        .Find "位置=" & lng位置
        If .EOF = False Then
            vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("状态")) = Decode(NVL(!执行状态), 3, "退料", "不处理")
            vsGrid.Cell(flexcpBackColor, lngRow, vsGrid.ColIndex("状态")) = Decode(NVL(!执行状态), 3, &HFFC0C0, &HFFFFFF)
            vsGrid.Cell(flexcpBackColor, lngRow, vsGrid.ColIndex("退料数量")) = Decode(NVL(!执行状态), 3, &HFFC0C0, &HFFFFFF)
            vsGrid.Cell(flexcpForeColor, lngRow, vsGrid.ColIndex("状态")) = Decode(NVL(!执行状态), 3, vbBlue, &H80000008)
            vsGrid.Cell(flexcpForeColor, lngRow, vsGrid.ColIndex("退料数量")) = Decode(NVL(!执行状态), 3, vbBlue, &H80000008)
        End If
        .MoveFirst
        .Find "执行状态=3"
        mblnHave退料 = (.EOF = False)
    End With
End Sub
Private Sub SetGRDCOLOR(ByVal objGrd As Object, ByVal lngRow As Long, ByVal int记录状态 As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置控件的显示颜色
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-03 17:11:56
    '-----------------------------------------------------------------------------------------------------------

    Dim lngColor As Long
    Dim i As Long
    If int记录状态 = 1 Then
        lngColor = &H80000008
    ElseIf zlCommFun.ZyMod(int记录状态, 3) = 2 Then
         lngColor = vbRed
    Else
        lngColor = vbBlue
    End If
    With vsGrid
        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = lngColor
    End With
End Sub

 Private Function 获取接收部门条件(ByVal strDate As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取接收部门的打印条件
    '入参:
    '出参:
    '返回:成功,返回 显示|IN(部门ID,..) ,否则返回""
    '编制:刘兴洪
    '日期:2008-05-05 13:31:28
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, str显示 As String, strIDIn As String
    If strDate = "" And mArrFilter("开单科室id") = "" Then Exit Function
    
    On Error GoTo ErrHandle
    If mArrFilter("开单科室id") = "" And strDate <> "" Then
        '没有条件,则以根据选择的类别读取显示科室
        gstrSQL = "Select distinct D.ID,D.编码,D.名称 as 科室 " & _
                 " From 药品收发记录 S,门诊费用记录 C,部门表 d " & _
                 " Where S.费用ID=C.ID And Mod(S.记录状态,3) In (0,1) And S.审核人 Is Not Null " & _
                 "      And C.执行状态=1 And S.库房ID=[1] And S.发药方式=3 And S.审核日期=[2] " & _
                 "      And S.单据 In (24,25,26) "
        Select Case Val(mArrFilter("部门类型"))
            Case 0  '
                gstrSQL = gstrSQL & " and C.病人科室id=d.id(+) "
            Case 1 '医技
                gstrSQL = gstrSQL & "  and C.开单部门id =d.id(+)"
            Case Else '病区
                gstrSQL = gstrSQL & "  and C.病人病区ID =d.id(+)"
        End Select
        
        If mArrFilter("单据") = "24" Then
            If Val(mArrFilter("部门类型")) = 2 Then Exit Function
        ElseIf mArrFilter("单据") = "26" Then
            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        ElseIf InStr(1, mArrFilter("单据"), "25") > 0 Or InStr(1, mArrFilter("单据"), "26") > 0 Then
            If InStr(1, mArrFilter("单据"), "24") > 0 And Val(mArrFilter("部门类型")) = 2 Then
                gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
            Else
                gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
            End If
        End If
        
        gstrSQL = gstrSQL & "order by 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mArrFilter("发料部门id")), CDate(strDate))
        With rsTemp
            Do While Not .EOF
                If NVL(!科室, "") <> "" Then
                    str显示 = str显示 & "," & !科室
                    strIDIn = strIDIn & "," & !Id
                End If
                
                rsTemp.MoveNext
            Loop
        End With
        
        If str显示 = "" Then
            获取接收部门条件 = "所有科室|Is Not Null"
        Else
            strIDIn = "0" & strIDIn
            str显示 = str显示 & "|" & " IN (" & strIDIn & ")"
            获取接收部门条件 = str显示
        End If
        
        Exit Function
    End If
    gstrSQL = "Select ID, 名称 From 部门表 A, Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) J Where ID = J.Column_Value order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(mArrFilter("开单科室id")))
    With rsTemp
        Do While Not .EOF
            str显示 = str显示 & "," & !名称
            rsTemp.MoveNext
        Loop
    End With
    str显示 = str显示 & "|" & " IN (" & CStr(mArrFilter("开单科室id")) & ")"
    获取接收部门条件 = str显示
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

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


Private Sub vsGrid_EnterCell()
    With vsGrid
        If .Row = 0 Then Exit Sub

        .Editable = flexEDNone
        .FocusRect = flexFocusLight

        If .Row > 0 Then
            Select Case .Col
            Case .ColIndex("退料数量")
                .Editable = flexEDKbdMouse
                .FocusRect = flexFocusSolid
            Case .ColIndex("状态")
                .FocusRect = flexFocusSolid
            End Select
        End If

    End With
End Sub


Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With vsGrid
        strKey = .EditText
        Select Case .Col
            Case .ColIndex("退料数量")
                intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.数量小数, g_小数位数.obj_散装小数.数量小数)
        End Select
        
        If Col = .ColIndex("退料数量") Then
            If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13) + Chr(Asc("-")), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
                Exit Sub
            ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                If .EditSelLength = Len(strKey) Then Exit Sub
                If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                    KeyAscii = 0
                    Exit Sub
                End If
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub


