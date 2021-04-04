VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalannceGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private WithEvents vsBlance As VSFlexGrid
Attribute vsBlance.VB_VarHelpID = -1
Private mfrmMain As Object
Private mlngModule As Long
Private mEditType As gBalanceBill
Private mobjImgDelIco As IPictureDisp

Public Event AfterEdit(ByVal Row As Long, ByVal Col As Long)
Public Event AfterDeleteRow()
Public Event AfterReCalcYBMoney(ByRef objInsure As clsInsureInfor)
Public Event AfterPayMoney(ByRef objItems As clsBalanceItems, ByVal blnSaved As Boolean, ByVal blnSucces As Boolean, ByVal blnChangeMoney As Boolean)
Public Event GetBalanceSaveSQL(ByVal objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, ByRef cllBillPro As Collection, blnCancel As Boolean)
Public Event GetSaveStrickDepositSQL(ByVal objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, ByRef cllBillPro As Collection, blnCancel As Boolean)

Public Event ShowAccountMoney(ByVal objCard As Card, ByVal dbl帐户余额 As Double)
Private mvarbln已报价 As Boolean '局部复制
Private mvarobjBalanceInfor As clsBalanceInfo '局部复制
Private mvarobjPatiInfor As clsPatientInfo  '局部复制
Private mvarblnBatchState As Boolean '局部复制
Private mvarbln医保未结算 As Boolean '局部复制
Private mobjDelCashItems As clsBalanceItems '强制退现
Private mobjThirdSwap As clsThirdSwap   '三方交易接口
Private mstrPrivsCard As String
Private mblnNotChange As Boolean
Private mbyt刷卡缺省金额操作 As Byte
Private mlng合约单位ID As Long
Private mstr费用性质 As String
Private mrsFeeList As ADODB.Recordset   '费用信息集，主要是为了计算消费卡
Private mrsBalance As ADODB.Recordset
Private mbytMode As Byte '结算模式:0-传统模式;1-缴款模式;2-合约单位结帐
'日志相关变量
Private mstrErrInfor As String, mstrFunName As String, mblnReturnValue As Boolean

Public Property Let bln医保未结算(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.bln医保未结算 = 5
    mvarbln医保未结算 = vData
End Property


Public Property Get bln医保未结算() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.bln医保未结算
    bln医保未结算 = mvarbln医保未结算
End Property



Public Property Let blnBatchState(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.blnBatchState = 5
    mvarblnBatchState = vData
End Property


Public Property Get blnBatchState() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.blnBatchState
    blnBatchState = mvarblnBatchState
End Property

Public Property Set objPatiInfor(ByVal vData As clsPatientInfo)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objPatiInfor = Form1
    Set mvarobjPatiInfor = vData
End Property


Public Property Get objPatiInfor() As clsPatientInfo
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objPatiInfor
    Set objPatiInfor = mvarobjPatiInfor
End Property
Public Property Set objBalanceInfor(ByVal vData As clsBalanceInfo)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objBalanceInfor = Form1
    Set mvarobjBalanceInfor = vData
End Property


Public Property Get objBalanceInfor() As clsBalanceInfo
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objBalanceInfor
    Set objBalanceInfor = mvarobjBalanceInfor
End Property

Public Property Let bln已报价(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.bln已报价 = 5
    mvarbln已报价 = vData
End Property


Public Property Get bln已报价() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.bln已报价
    bln已报价 = mvarbln已报价
End Property

Public Function InitComm(ByRef frmMain As Object, ByRef objImgDelIco As IPictureDisp, ByRef objThirdSwap As clsThirdSwap, _
    ByRef objBalanceGrid As VSFlexGrid, ByVal lngModule As Long, ByVal EditType As gBalanceBill, Optional bytMode As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化
    '入参:frmMain-有界面
    '     objBalanceGrid-网格控制
    '     lngModule-模块号
    '     EditType-编辑类型
    '     objImgDelIco-删除图标
    '     bytMode- 结算模式:0-传统模式;1-缴款模式;2-合约单位结帐
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-03 13:26:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain: Set vsBlance = objBalanceGrid
    mEditType = EditType: mlngModule = lngModule
    Set mobjImgDelIco = objImgDelIco
    Set mobjDelCashItems = New clsBalanceItems
    Set mobjThirdSwap = objThirdSwap
    mstrPrivsCard = ";" & GetPrivFunc(glngSys, 1151) & ";"
    mbyt刷卡缺省金额操作 = Val(zlDatabase.GetPara("刷卡缺省金额操作", glngSys, 1151, "0")) '86853
    
    mbytMode = bytMode
    InitComm = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReInitVar(ByRef objPati_In As clsPatientInfo, objBalanceInfor_In As clsBalanceInfo, rsFeeList As ADODB.Recordset, rsBalance As ADODB.Recordset, _
    lng合约单位ID As Long, str费用性质 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化模块相关变量
    '入参:str费用性质-消费卡使用时才有效
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-03 14:51:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set objPatiInfor = objPati_In: Set objBalanceInfor = objBalanceInfor_In: mlng合约单位ID = lng合约单位ID: mstr费用性质 = str费用性质
    Set mrsFeeList = rsFeeList
    Set mrsBalance = mrsBalance
    zlReInitVar = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    Set mobjDelCashItems = New clsBalanceItems
    Set mvarobjBalanceInfor = New clsBalanceInfo
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    zl_vsGrid_Para_Restore mlngModule, vsBlance, mfrmMain.Name, "结算列表"
    Set mobjDelCashItems = Nothing
    Set mobjImgDelIco = Nothing
    Set mvarobjBalanceInfor = Nothing
    Set mvarobjPatiInfor = Nothing
End Sub

Private Sub vsBlance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblCurMoney As Double, dbl个人帐户 As Double, dbl医保基金 As Double, dbl脱机医保 As Double
    Dim objCard As Card, objItem As clsBalanceItem
    With vsBlance
        Select Case Col
        Case .ColIndex("结算方式")
        Case .ColIndex("结算金额")
            If zlGetBalanceItemFromBalanceGrid(vsBlance, Row, objItem) = False Then Exit Sub
            If objItem.是否脱机医保 Then bln已报价 = False
            
            objBalanceInfor.医保支付合计 = RoundEx(GetYBTotal(0, dbl个人帐户, dbl医保基金, dbl脱机医保), 5)
            objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.医保支付合计 + dbl脱机医保, 5)
            objBalanceInfor.未付合计 = RoundEx(objBalanceInfor.当前结帐 - objBalanceInfor.已付合计, 5)
            RaiseEvent AfterEdit(Row, Col)
            bln已报价 = False
        Case Else
        End Select
    End With
End Sub

Private Sub vsBlance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBlance, mfrmMain.Name, "结算列表"
End Sub

Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vsBlance, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsBlance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBlance, mfrmMain.Name, "结算列表"
End Sub

Private Sub vsBlance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim objPati As clsPatientInfo, varTemp As Variant
    Dim objItem As clsBalanceItem
    
    If mEditType = g_Ed_单据查看 Or blnBatchState Or bln医保未结算 Then Cancel = True: Exit Sub    '医保未结算的，不允许编辑
      
    If GetPatiObject(objPati) = False Then Cancel = True: Exit Sub
    If objPati.病人ID = 0 Then Cancel = True: Exit Sub
    
    
    With vsBlance
        If Val(.TextMatrix(Row, .ColIndex("结算状态"))) = 1 Then Cancel = True: Exit Sub
        
        varTemp = Split(.TextMatrix(Row, .ColIndex("编辑状态")) & "|||", "|")
        .ComboList = ""
        Select Case Col
        Case .ColIndex("结算方式")
            If Not zlGetBalanceItemFromBalanceGrid(vsBlance, Row, objItem) Then Cancel = True: Exit Sub
            If Val(varTemp(1)) <> 1 Then Cancel = True: Exit Sub
            If objItem.是否脱机医保 Then Cancel = True: Exit Sub
            '是否允许编辑|是否允许删除
             .ColComboList(.ColIndex("结算方式")) = ""
             .ComboList = "..."
             .CellButtonPicture = mobjImgDelIco
        Case .ColIndex("结算号码"), .ColIndex("备注")
            If Val(varTemp(0)) <> 1 Then Cancel = True: Exit Sub
            If Val(.TextMatrix(Row, .ColIndex("结算性质"))) <> 2 Then Cancel = True: Exit Sub
        Case .ColIndex("结算金额")
            If Val(varTemp(0)) <> 1 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub

Private Sub DeletePayInfor(ByVal lngDelRow As Long, Optional ByVal blnForceDel As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除支付信息
    '编制:刘兴洪
    '日期:2015-01-28 15:18:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, objItemTemp As clsBalanceItem, objItems As clsBalanceItems
    Dim lngRow As Long, str操作员姓名 As String
    Dim str结算信息 As String
    On Error GoTo errHandle
    
    If mobjDelCashItems Is Nothing Then Set mobjDelCashItems = New clsBalanceItems
     
    With vsBlance
        If lngDelRow > .Rows - 1 Or lngDelRow < 1 Then Exit Sub
        If zlGetBalanceItemFromBalanceGrid(vsBlance, lngDelRow, objItem) = False Then Exit Sub
         Select Case objItem.结算类型   '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
         Case 3 '
            If objItem.是否退款 Then    '退款退现检查
            
                If objItem.是否保存 Then
                    '已经保存了的，需要判断交易是否正在进行中，如果正在进行中，则不允许删除
                    If zlGetBalanceItemsFromVsBalanceGrid(vsBlance, objItem, objItems) = False Then Exit Sub
                    If mobjThirdSwap.zlThird_IsCancelFromItems(objItems) = False Then Exit Sub
                    'If mobjThirdSwap.zlDelBalanceItemFromItemObject(objItem) = False Then Exit Sub '删除数据
                End If
                
                If Not objItem.是否允许退现 And objItem.结算类型 = 3 And objItem.结算金额 <> 0 And (objItem.是否结算 Or mEditType = g_Ed_结帐作废 Or objItem.是否预交) Then
                    If zlGetBalanceItemsFromVsBalanceGrid(vsBlance, objItem, objItems, False, False, str结算信息) = False Then Exit Sub
                    If InStr(";" & mstrPrivsCard & ";", ";三方退款强制退现;") = 0 And objItem.是否强制退现 Then
                        '已经验证过的，不再验证
                        str操作员姓名 = zlDatabase.UserIdentifyByUser(Me, "强制退现验证", glngSys, 1151, "三方退款强制退现")
                        If str操作员姓名 = "" Then
                            MsgBox "录入的操作员验证失败或者录入的操作员不具备强制退现权限，不能强制退现！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        For Each objItemTemp In objItems
                            objItemTemp.Tag = str操作员姓名
                            mobjDelCashItems.AddItem objItemTemp
                        Next
                    Else
                        If InStr(str结算信息, "|") > 0 Then
                            If MsgBox("以下结算不支持退现,是否强制退现？" & vbCrLf & Replace(str结算信息, "|", vbCrLf), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
                        Else
                            If MsgBox(.TextMatrix(lngDelRow, .ColIndex("结算方式")) & "不支持退现,是否强制退现？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
                        End If
                        
                        For Each objItemTemp In objItems
                            objItemTemp.Tag = str操作员姓名
                            mobjDelCashItems.AddItem objItemTemp
                        Next
                    End If
                    If objItem.是否保存 Then
                        If mobjThirdSwap.zlDelBalanceItemFromItemObject(objItem) = False Then Exit Sub '删除数据
                    End If
                End If
                
                lngRow = lngDelRow
                '删除指定的行:
                Call zlMoveRowBalanceFromSwapID(vsBlance, objItem.卡类别ID, False, objItem.关联交易ID, objItem.预交ID)
                Call zlReCalcBalanceInfor(vsBlance, objBalanceInfor)
                Call zlRecalItemObjectRowNo(vsBlance)   '重新刷行对象的行号
                If .RowIsVisible(.Row) = False Then .ShowCell .Row, .Col
                bln已报价 = False
                RaiseEvent AfterDeleteRow
                Exit Sub
            ElseIf objItem.是否保存 Then     '扣款交易
                If zlGetBalanceItemsFromVsBalanceGrid(vsBlance, objItem, objItems) = False Then Exit Sub
                If mobjThirdSwap.zlThird_IsCancelFromItems(objItems) = False Then Exit Sub
                If mobjThirdSwap.zlDelBalanceItemFromItemObject(objItem) = False Then Exit Sub '删除数据
            End If
         Case Else
         End Select
        
        lngRow = lngDelRow
        'Call LoadCurOwnerPayInfor
        
        If .Rows <= 2 Then
            .Clear 1
            .RowData(1) = "": Set objItem = Nothing
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        Else
            .RowData(lngDelRow) = ""
            Set objItem = Nothing
            vsBlance.RemoveItem lngDelRow
        End If
        
        If lngRow <= 1 Then
            lngRow = 1
        ElseIf lngRow >= .Rows - 1 Then
            lngRow = .Rows - 1
        Else
            lngRow = lngDelRow + 1
        End If
        If lngRow > .Rows - 1 Or lngRow <= 1 Then lngRow = 1
        .Row = lngRow
        If .RowIsVisible(.Row) = False Then .ShowCell .Row, .Col
        Call zlReCalcBalanceInfor(vsBlance, objBalanceInfor)
    End With
    Call zlRecalItemObjectRowNo(vsBlance)   '重新刷行对象的行号
    bln已报价 = False
    
    RaiseEvent AfterDeleteRow
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub vsBlance_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim varData As Variant
    If mEditType = g_Ed_单据查看 Then Exit Sub
    With vsBlance
        '是否允许编辑|是否允许删除
        varData = Split(.TextMatrix(Row, .ColIndex("编辑状态")) & "||", "|")
        If varData(1) <> 1 Then Exit Sub
    End With
    
    Call DeletePayInfor(Row)
End Sub

 
Private Sub vsBlance_ChangeEdit()
    Dim objItem As clsBalanceItem
    Dim dblMoney As Double
    If mEditType = g_Ed_单据查看 Then Exit Sub
    With vsBlance
       Select Case .Col
       Case .ColIndex("结算金额")
'           Call zlGetBalanceItemFromBalanceGrid(vsBlance, .Row, objItem)
'           If objItem Is Nothing Then Exit Sub
'           dblMoney = objItem.结算金额 - RoundEx(Val(.EditText), 2)
     
       Case Else
       End Select
    End With
End Sub


Private Sub vsBlance_DblClick()
    Dim varData As Variant
    If mEditType = g_Ed_单据查看 Then Exit Sub
    With vsBlance
        If .Col <> .ColIndex("结算金额") Then Exit Sub
        '是否允许编辑|是否允许删除
        varData = Split(.TextMatrix(.Row, .ColIndex("编辑状态")) & "||", "|")
        If Val(varData(0)) <> 1 Then Exit Sub
        .EditCell
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub
 
Private Sub vsBlance_EnterCell()
    Dim varData As Variant
    Dim objItem As clsBalanceItem
    Dim dblMoney As Double
    Dim lngRow As Long
    
    If mEditType = g_Ed_单据查看 Then Exit Sub
   
    With vsBlance
        If .Row < 0 Then Exit Sub
        
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        Select Case Val(.TextMatrix(.Row, .ColIndex("结算性质")))
        Case 2
            .ColData(.ColIndex("结算号码")) = "0||0"
            .ColData(.ColIndex("备注")) = "0||0"
        Case Else
            .ColData(.ColIndex("结算号码")) = "0||2"
            .ColData(.ColIndex("备注")) = "0||2"
        End Select
    End With
End Sub


Private Sub vsBlance_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim varTemp As Variant, lngRow As Long

    If mEditType = g_Ed_单据查看 Then Exit Sub
    
    With vsBlance
        If .Row > .Rows - 1 Or .Row < 1 Then Exit Sub
        varTemp = Split(.TextMatrix(.Row, .ColIndex("编辑状态")) & "||||", "|")
        
        If KeyCode <> vbKeyReturn And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) And Val(varTemp(1)) = 1 Then
                Call vsBlance_CellButtonClick(.Row, .Col)
                Exit Sub
            End If
        End If
        '删除
        If KeyCode = vbKeyDelete Then
            Call vsBlance_CellButtonClick(.Row, .Col)
            Exit Sub
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    mblnNotChange = True
    With vsBlance
        Select Case .Col
        Case .ColIndex("结算方式")
            If Trim(.TextMatrix(.Row, .ColIndex("结算方式"))) = "" And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        Case .ColIndex("结算金额")
            If (Trim(.TextMatrix(.Row, .ColIndex("结算方式"))) = "" Or Val(.TextMatrix(.Row, .ColIndex("结算金额"))) = 0) And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        Case Else
            If (Trim(.TextMatrix(.Row, .ColIndex("结算方式"))) = "" Or Val(.TextMatrix(.Row, .ColIndex("结算金额"))) = 0) And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        End Select
        Call zlVsMoveGridCell(vsBlance, .ColIndex("结算方式"), , IIf(mEditType = g_Ed_单据查看 Or mEditType = g_Ed_结帐作废, False, True), lngRow)
    End With
    mblnNotChange = False
    Call vsBlance_EnterCell
End Sub

Private Sub vsBlance_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strKey As String, lngRow As Long
    
    If mEditType = g_Ed_单据查看 Or mEditType = g_Ed_结帐作废 Then Exit Sub
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsBlance
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '暂不处理输入
        Select Case Col
        Case .ColIndex("结算方式")
        Case .ColIndex("结算金额")
        Case Else
        End Select
        Call zlVsMoveGridCell(vsBlance, .ColIndex("结算方式"), -1, True, lngRow)
    End With

End Sub

Private Sub vsBlance_KeyPress(KeyAscii As Integer)
    Dim varTemp As Variant
    
    With vsBlance
        If .Row > .Rows - 1 Or .Row <= 0 Or .Col > .Cols - 1 Or .Col < 0 Then Exit Sub
        varTemp = Split(.TextMatrix(.Row, .ColIndex("编辑状态")) & "|", "|")
        If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If Not (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Or mEditType = g_Ed_重新结帐) Then KeyAscii = 0: Exit Sub
        '是否允许编辑|是否允许删除
        varTemp = Split(.TextMatrix(.Row, .ColIndex("编辑状态")) & "|", "|")
        If Val(varTemp(0)) <> 1 Then KeyAscii = 0: Exit Sub
        Select Case .Col
        Case .ColIndex("结算金额")
            Call VsFlxGridCheckKeyPress(vsBlance, vsBlance.Row, vsBlance.Col, KeyAscii, m金额式)
        End Select
    End With
End Sub

Private Sub vsBlance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

  If Not (mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Or mEditType = g_Ed_重新结帐) Then Exit Sub
  
  With vsBlance
        Select Case .Col
        Case .ColIndex("结算金额")
            Call VsFlxGridCheckKeyPress(vsBlance, Row, Col, KeyAscii, m负金额式)
        Case .ColIndex("结算号码"), .ColIndex("备注")
            Call VsFlxGridCheckKeyPress(vsBlance, Row, Col, KeyAscii, m文本式)
            Exit Sub
        Case Else
            KeyAscii = 0: Exit Sub
        End Select
    End With
End Sub

Private Sub vsBlance_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strKey As String, lngRow As Long
    If mEditType = g_Ed_单据查看 Or mEditType = g_Ed_结帐作废 Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsBlance
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '暂不处理输入
        Select Case Col
        Case .ColIndex("结算方式")
        Case .ColIndex("结算金额")
        Case Else
        End Select
       ' Call zlVsMoveGridCell(vsBlance, .ColIndex("结算方式"), -1, True, Row)
    End With
End Sub

Private Sub vsBlance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Err = 0: On Error GoTo ErrHand:
    With vsBlance
        If .MouseRow < 1 Or .MouseRow > .Rows - 1 Then Exit Sub
        If .MouseCol < 0 Or .MouseCol > .Cols - 1 Then Exit Sub
        
       If .ToolTipText = Trim(.TextMatrix(.MouseRow, .MouseCol)) Then Exit Sub
       .ToolTipText = Trim(.TextMatrix(.MouseRow, .MouseCol))
    End With
ErrHand:
    Exit Sub
End Sub

Private Sub vsBlance_LeaveCell()
    If mEditType = g_Ed_单据查看 Then Exit Sub
    OS.OpenIme False
End Sub
Private Sub vsBlance_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If mEditType = g_Ed_单据查看 Then Exit Sub
    '设置单元格的编辑长度
    With vsBlance
       Select Case .Col
           Case .ColIndex("结算方式")
               .EditMaxLength = 50
           Case .ColIndex("结算金额")
               .EditMaxLength = 16
           Case .ColIndex("结算号码")
               .EditMaxLength = 30
           Case .ColIndex("备注")
               .EditMaxLength = 50
           Case Else
               .EditMaxLength = 100
       End Select
    End With
End Sub
Private Sub vsBlance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim objCard As Card, dbl原始金额 As Double, dblTotal As Double
    Dim str结算方式 As String
    Dim dblMoney As Double, blnYb As Boolean
    Dim strInput As String
    Dim objPati As clsPatientInfo, objItem As clsBalanceItem
    Dim objInsureInfor As clsInsureInfor
    
    
    Call GetPatiObject(objPati)
    
    objInsureInfor = objBalanceInfor.objInsure
    
    With vsBlance
        If Row <= 0 Then Exit Sub
         strInput = Trim(.EditText): strInput = Replace(strInput, Chr(vbKeyReturn), ""): strInput = Replace(strInput, Chr(10), "")
      
        Select Case Col
        Case .ColIndex("结算方式")
        
      
        Case .ColIndex("结算金额")
            
            If Not IsNumeric(strInput) And strInput <> "" Then
                MsgBox "输入的金额必须为数字！", vbInformation, gstrSysName
                .EditCell: .EditSelStart = 0
                .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True
                Exit Sub
            End If
            
            If zlDblIsValid(strInput, 16, False, False, 0, .ColKey(Col)) = False Then
                Cancel = True: Exit Sub
            End If
            If zlGetBalanceItemFromBalanceGrid(vsBlance, Row, objItem) = False Then Cancel = True: Exit Sub
            If objItem.是否脱机医保 Then
                strInput = Format(Val(strInput), "0.00")
                .EditText = strInput
                .TextMatrix(Row, Col) = strInput
                Exit Sub
            End If
        
            If Val(strInput) < 0 Then
                MsgBox "输入的金额必须为大于等于零！", vbInformation, gstrSysName
                .EditCell: .EditSelStart = 0
                .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True
                Exit Sub
            End If
            
            str结算方式 = Trim(.TextMatrix(.Row, .ColIndex("结算方式")))
            If str结算方式 = "" Then Exit Sub
                
            '结算金额不允许超过返回的原始金额(个人帐户允许透支时再判断)
            dbl原始金额 = Val(.Cell(flexcpData, .Row, .ColIndex("结算金额")))
            Select Case Val(.TextMatrix(.Row, .ColIndex("结算性质")))
            Case 3 '个人帐户
                If Val(strInput) > dbl原始金额 And Val(strInput) <> 0 And dbl原始金额 <> 0 Then
                    MsgBox "输入的""" & str结算方式 & """结算金额不能超过 " & Format(dbl原始金额, "0.00") & " ！", vbInformation, gstrSysName
                    .EditCell: .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                 '不允许超过允许透支金额
                If objInsureInfor.帐户余额 + objInsureInfor.个帐透支 - Val(strInput) < 0 Then
                    MsgBox "帐户余额:" & Format(objInsureInfor.帐户余额, "0.00") & _
                        IIf(objInsureInfor.个帐透支 = 0, "", "(" & "允许透支:" & Format(objInsureInfor.个帐透支, "0.00") & ")") & _
                        "不足要结算的金额。", vbInformation, gstrSysName
                    .EditCell: .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                blnYb = True
            Case 4 '医保基金
                If Val(strInput) > dbl原始金额 And Val(strInput) <> 0 And dbl原始金额 <> 0 Then
                    MsgBox "输入的""" & str结算方式 & """结算金额不能超过 " & Format(dbl原始金额, "0.00") & " ！", vbInformation, gstrSysName
                    .EditCell: .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                blnYb = True
            End Select
                
            '不允许超出单据剩余可结算金额
            dblTotal = RoundEx(objBalanceInfor.当前结帐 - GetYBTotal(.Row) - objBalanceInfor.已付合计, 5)
            If Val(strInput) > dblTotal And RoundEx(Val(strInput), 6) <> 0 And blnYb = False Then
                MsgBox "结算金额过大，超过单据允许结算金额:" & Format(dblTotal, "0.00") & "。", vbInformation, gstrSysName
                .EditCell
                .EditSelStart = 0
                .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True
                Exit Sub
            End If
        
            strInput = Format(Val(strInput), "0.00")
            .EditText = strInput
            .TextMatrix(Row, Col) = strInput
            '重新计算医保结算金额
            If blnYb Then Call ReCalcYBMoney
         Case .ColIndex("结算号码"), .ColIndex("备注")
            If strInput = "" Then Exit Sub
            If zlCommFun.StrIsValid(strInput, .EditMaxLength, , .ColKey(Col)) = False Then Cancel = True
        Case Else
        End Select
    End With
End Sub


Private Function GetYBTotal(ByVal lngRow As Long, Optional ByRef dbl个人帐户 As Double, Optional ByRef dbl医保基金 As Double, Optional ByRef dbl脱机医保合计 As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医保支付总额
    '入参:lngRow-不包含的行
    '出参:dbl个人帐户-返回个人账户
    '     dbl医保基金-返回医保金额
    '     dbl脱机医保合计-获取脱机医保合计
    '返回:医保支付总额
    '编制:刘兴洪
    '日期:2015-01-21 15:41:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblMoney As Double, str结算方式 As String
    Dim objItem As clsBalanceItem
    On Error GoTo errHandle
    dbl脱机医保合计 = 0
    With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(.Row, .ColIndex("结算方式")))
            If str结算方式 <> "" And i <> lngRow Then
                If zlGetBalanceItemFromBalanceGrid(vsBlance, i, objItem) Then
                     If objItem.卡类别ID = 0 Then
                        If Not objItem.是否脱机医保 Then
                            '结算性质:结算方式.性质
                             dblMoney = Val(.TextMatrix(i, .ColIndex("结算金额")))
                             Select Case objItem.结算性质
                             Case 3 '个人帐户
                                dbl个人帐户 = RoundEx(dbl个人帐户 + dblMoney, 6)
                             Case 4 '医保基金
                                dbl医保基金 = RoundEx(dbl医保基金 + dblMoney, 6)
                             End Select
                        Else
                            dbl脱机医保合计 = RoundEx(dbl脱机医保合计 + objItem.结算金额, 6)
                        End If
                    End If
                End If
            End If
        Next
    End With
    GetYBTotal = dbl个人帐户 + dbl医保基金
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Function GetPatiObject(ByRef objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息对像集
    '入参:
    '出参:objPati-病人信息对象集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-29 11:04:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Set objPati = objPatiInfor
    If objPati Is Nothing Then Set objPati = New clsPatientInfo: Set objPatiInfor = objPati
    GetPatiObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ReCalcYBMoney(Optional blnOlnyCalcYB As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算医保金额
    '入参：blnOlnyCalcYB-仅计算医保
    '编制:刘兴洪
    '日期:2015-01-21 15:41:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, i As Long
    Dim dbl个人帐户 As Double, dbl医保基金 As Double, dblMoney As Double
    Dim str结算方式 As String
    
    On Error GoTo errHandle
    With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            If str结算方式 <> "" Then
                If zlGetBalanceItemFromBalanceGrid(vsBlance, i, objItem) Then
                    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                    If objItem.卡类别ID = 0 Then
                         dblMoney = Val(.TextMatrix(i, .ColIndex("结算金额")))
                         Select Case objItem.结算性质
                         Case 3 '个人帐户
                            dbl个人帐户 = dbl个人帐户 + dblMoney
                         Case 4 '医保基金
                            dbl医保基金 = dbl医保基金 + dblMoney
                         End Select
                    End If
                End If
            End If
        Next
    End With
        
    objBalanceInfor.医保支付合计 = RoundEx(dbl个人帐户 + dbl医保基金, 5)
    objBalanceInfor.objInsure.个帐支付 = dbl个人帐户
    objBalanceInfor.objInsure.统筹支付 = dbl医保基金
    If blnOlnyCalcYB Then Exit Sub
    
    RaiseEvent AfterReCalcYBMoney(objBalanceInfor.objInsure)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function Excute_Third_NotPlaySuccesBalances(ByVal objPati As clsPatientInfo, objBalanceInfor As clsBalanceInfo, _
    Optional blnChangeMoney_Out As Boolean, Optional blnSaveed_Out As Boolean, Optional objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行交易未成功的扣款数据
    '入参:objPati-当前病人对象
    '     objBalanceInfor-已经结帐的结帐数据
    '出参:blnChangeMoney_Out-是否金额发生了改变
    '     objBalanceInfor-返回最新的结帐数据(更新后的)
    '     blnSaveed_Out-是否保存了数据
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-14 10:35:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItems As clsBalanceItems, i As Long, objItem As clsBalanceItem, blnChangeMoneyTemp As Boolean
    On Error GoTo errHandle
ReDO:
    Call zlReCalcBalanceInfor(vsBlance, objBalanceInfor)
    Call zlRecalItemObjectRowNo(vsBlance)     '重新计算结算信息对象的行号值
    
    blnChangeMoney_Out = False
    With vsBlance
        For i = 1 To .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsBlance, i, objItem) Then
                If objItem.卡类别ID <> 0 And Not objItem.是否退款 And objItem.是否结算 = False And objItem.消费卡 = False And objItem.结算金额 <> 0 Then
                    '不是退款，直接扣款
                    '一卡通接口支付
                    If zlGetBalanceItemsFromVsBalanceGrid(vsBlance, objItem, objItems) = False Then Exit Function
                    objBalanceInfor.已付合计 = objBalanceInfor.已付合计 - objItems.结算金额
                    objBalanceInfor.未付合计 = objBalanceInfor.未付合计 + objItems.结算金额
                    If Excute_Third_Payment(objPati, objItems, objBalanceInfor, blnSaveed_Out, objSetFocus, blnChangeMoneyTemp) = False Then
                        Call zlReCalcBalanceInfor(vsBlance, objBalanceInfor)
                        Exit Function
                    End If
                    If blnChangeMoneyTemp Then blnChangeMoney_Out = True
                    GoTo ReDO: '可以结算列表发生列变化
                End If
            End If
        Next
    End With
    
    Excute_Third_NotPlaySuccesBalances = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Check_ThirdPayment_IsValied(ByVal objCurItem As clsBalanceItem, ByVal objPati As clsPatientInfo, _
    ByRef objBalanceInfo As clsBalanceInfo, ByRef objBalanceItems_out As clsBalanceItems, Optional blnCurInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方交易支付验证
    '入参:objCurItem-当前结算信息
    '     objBalanceInfo-当前结帐信息
    '     blnCurInput-是否当前输入项
    '出参:objBalanceItems_Out-当前刷卡结算信息
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '说明:同步验证了接口和刷卡接口的
    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim dblTemp As Double, cllSquareBalance As Collection, objDelBalanceItems As clsBalanceItems
'    Dim strXMLExpend As String, bln退现 As Boolean
'    Dim dbl帐户余额 As Double,
'    Dim strExpand As String, strXmlIn As String, strErrMsg As String
'    Dim strBalanceIDs As String
    Dim intMousePointer As Integer
    Dim dblMoney As Double, objCard As Card, dbl未付金额 As Double, dbl帐户余额 As Double
    Dim strFunName As String
    
    
     On Error GoTo errHandle
     
    strFunName = "Check_ThirdPayment_IsValied"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "objCurItem=" & IIf(objCurItem Is Nothing, "Nothing", " Not Nothing"))
    
    
    intMousePointer = Screen.MousePointer
    
    If objCurItem Is Nothing Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,objCurItem= Nothing")
        Exit Function
    End If
    Set objCard = objCurItem.objCard
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then
        Check_ThirdPayment_IsValied = True
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true,objCard.接口序号=" & objCard.接口序号 & ", objCard.消费卡= " & objCard.消费卡)
        Exit Function
    End If
    If zlCheckBalancesIsExistFromCardTypeID(vsBlance, objCard.接口序号, False) = False Then
        '暂无限制
    End If
    
   
     
    dbl未付金额 = RoundEx(objBalanceInfo.未付合计 - objBalanceInfo.冲预交合计 - objBalanceInfo.误差费, 6)
     
    dblMoney = objCurItem.结算金额
    If dblMoney = 0 Then
        If dbl未付金额 = 0 Then
            Check_ThirdPayment_IsValied = True
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true,dbl未付金额=" & dbl未付金额)
            Exit Function
        End If
        Screen.MousePointer = 0
        MsgBox "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,dblMoney= 0")
        Exit Function
    End If
         
    If Abs(dblMoney) > Format(Abs(dbl未付金额), "0.00") Then
        Screen.MousePointer = 0
        MsgBox "刷卡金额不能大于本次未付金额:" & Format(Abs(dbl未付金额), "0.00") & " ！", vbInformation, gstrSysName
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,dblMoney= " & Abs(dblMoney) & ",dbl未付金额=" & Format(Abs(dbl未付金额), "0.00"))
        Exit Function
    End If
    
    
    If dblMoney <> Format(dbl未付金额, "0.00") And blnCurInput Then
        Select Case mbyt刷卡缺省金额操作
        Case 1  '提示
            If MsgBox("刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,选择为否,刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，是否继续？")
                Exit Function
            End If
        Case 2
            MsgBox "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！", vbInformation, gstrSysName
            
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false, 刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！")
            Exit Function
        End Select
    End If
    
    
    If mobjThirdSwap.zlThird_Payment_IsValied(objPati, objBalanceInfo, objCurItem, objBalanceItems_out, dbl帐户余额, 0) = False Then
        RaiseEvent ShowAccountMoney(objCard, dbl帐户余额)
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,mobjThirdSwap.zlThird_Payment_IsValied=false")
        Exit Function
    End If
    RaiseEvent ShowAccountMoney(objCard, dbl帐户余额)
    
    If dbl帐户余额 <> 0 And dbl帐户余额 < dblMoney Then
        Screen.MousePointer = 0
        MsgBox objCard.名称 & "的帐户余额不足,不允许刷卡消费!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,dblMoney= " & dblMoney & ",dbl帐户余额=" & dbl帐户余额)
        Exit Function
    End If
    Check_ThirdPayment_IsValied = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true")
    Exit Function
errHandle:
    Screen.MousePointer = 0
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,Err=" & Err.Description)
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function
Public Function Check_ThirdReturnMoney_IsValied(ByVal objCard As Card, ByVal objPati As clsPatientInfo, _
    ByVal objBalanceInfor As clsBalanceInfo, ByRef objCurDelBalanceItems As clsBalanceItems, _
    ByRef objBalanceItems_out As clsBalanceItems, Optional blnCurInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方交易退款检查
    '入参:objCard-三方卡
    '     objCurDelBalanceItems-当前退款信息
    '     objBalanceInfor-结算信息
    '     blnCurInput-是否为当前输入退款检查
    '出参:objBalanceItems_Out-当前退款信息
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItems As clsBalanceItems
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer
    On Error GoTo errHandle
    '操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
    If mobjThirdSwap.zlThird_ReturnMoney_IsValied(objPati, objCard, 2, objBalanceInfor, objCurDelBalanceItems, objItems, blnCurInput) = False Then Exit Function
    Set objBalanceItems_out = objCurDelBalanceItems
    Check_ThirdReturnMoney_IsValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function
Public Function Check_ThirdTransferAccount_IsValied(ByVal objCard As Card, ByVal objPati As clsPatientInfo, _
    ByVal objBalanceInfor As clsBalanceInfo, ByRef objCurTransBalanceItems As clsBalanceItems, _
    ByRef objBalanceItems_out As clsBalanceItems, Optional blnCurInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方交易退款检查
    '入参:objCard-三方卡
    '     objCurTransBalanceItems-当前退款信息
    '     objBalanceInfor-结算信息
    '     blnCurInput-是否当前输入
    '出参:objBalanceItems_Out-当前退款信息
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    Dim strXMLExpend As String, bln退现 As Boolean
    Dim dbl帐户余额 As Double, dbl未退金额 As Double
    Dim strExpand As String, strXmlIn As String, strErrMsg As String
    Dim strBalanceIDs As String
    Dim intMousePointer As Integer
    Dim dblMoney As Double, dblTotal As Double
    
    
    intMousePointer = Screen.MousePointer
    On Error GoTo errHandle
    
     
    If objCard Is Nothing Then
        If InStr(";" & mstrPrivsCard & ";", ";三方接口消费;") = 0 Then
            MsgBox "你没有三方接口消费权限，无法调用接口部件！", vbInformation, gstrSysName
            Exit Function
        End If
        MsgBox "未找到退款接口,请检查接口部件！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then Check_ThirdTransferAccount_IsValied = True: Exit Function
    
    If objCurTransBalanceItems Is Nothing Then
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    If objCurTransBalanceItems.Count = 0 Then
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    dblMoney = objCurTransBalanceItems.结算金额
    dbl未退金额 = RoundEx(objBalanceInfor.未付合计 - objBalanceInfor.冲预交合计, 6)
    
    If dblMoney = 0 Then
        If dbl未退金额 = 0 Then Check_ThirdTransferAccount_IsValied = True: Exit Function
        Screen.MousePointer = 0
        MsgBox "退款金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If dbl未退金额 = 0 Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & objCard.名称 & "已经退完，不能再退！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    If Abs(dblMoney) > Format(Abs(dbl未退金额), "0.00") And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "退款金额不能大于本次未退金额:" & Format(Abs(dbl未退金额), "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
   
    
    If dblMoney <> dbl未退金额 And blnCurInput Then
        Select Case mbyt刷卡缺省金额操作
        Case 1 '提示
            Screen.MousePointer = 0
            If MsgBox("当前退款金额与未退金额不符,是否继续？" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(dblMoney, "0.00"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        Case 2 '禁止
            Call MsgBox("当前退款金额与未退金额不符,不能继续!" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(dblMoney, "0.00"), vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End Select
    End If
    
    Set objBalanceItems_out = New clsBalanceItems
    '转帐及代扣处理
    If mobjThirdSwap.zlThird_TransferAccounts_IsValied(objPati, objCard, objBalanceInfor, objCurTransBalanceItems, objBalanceItems_out, dbl帐户余额, strErrMsg) = False Then
        RaiseEvent ShowAccountMoney(objCard, dbl帐户余额)
        Screen.MousePointer = 0
        If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    RaiseEvent ShowAccountMoney(objCard, dbl帐户余额)
    objBalanceItems_out.是否转帐 = True
    Check_ThirdTransferAccount_IsValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCheckMulitInterfaceNumValied(ByVal objCard As Card, Optional bln预交 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是正同时存在两种以上接口(不含两种)
    '返回:不含两种以上接口的,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer, i As Long, int性质 As Integer, str结算方式 As String
    Dim varData As Variant, strErrMsg As String
    Dim intMousePointer As Integer
    On Error GoTo errHandle
    strErrMsg = ""
    intMousePointer = Screen.MousePointer
    
    If bln预交 Or objCard.接口序号 <= 0 Then zlCheckMulitInterfaceNumValied = True:        Exit Function
    With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
            If InStr("34", int性质) > 0 Then
                intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & str结算方式 & ":" & .TextMatrix(i, .ColIndex("结算金额"))
            End If
        Next
    End With
    
    If mbytMode = 2 Then    '合约单位结帐
        If intCount > 1 Then
            Screen.MousePointer = 0
            Call MsgBox("注意:" & vbCrLf & "   合约单位结帐目前只支持二种接口,现在已经存在如下接口交易:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    zlCheckMulitInterfaceNumValied = True
    
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
End Function
Public Function Check_Square_Playment_IsValied(ByVal objPati As clsPatientInfo, ByVal objCurItem As clsBalanceItem, _
    ByRef objBalanceInfo As clsBalanceInfo, ByRef objBalanceItems_out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡结算交易检查
    '入参:objCurItem-当前结算信息
    '     objBalanceInfo-当前结帐数据
    '出参:objBalanceItems-当前刷卡结帐数据
    '     dblMoney-当前刷卡金额
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '说明:同步验证了接口和刷卡接口的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dbl未付金额 As Double, dbl帐户余额 As Double
    Dim rsMoney As ADODB.Recordset
    Dim objBalanceItemsOld As clsBalanceItems   '已经刷的消费卡结算信息
    Dim objCard As Card
    Dim intMousePointer As Integer
    Dim strFunName As String, str费用性质 As String
    intMousePointer = Screen.MousePointer

    strFunName = "Check_Square_Playment_IsValied"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "objCurItem=" & IIf(objCurItem Is Nothing, "Nothing", " Not Nothing"))
    
    If objCurItem Is Nothing Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,objCurItem=Nothing")
        Exit Function
    End If
    
    Set objCard = objCurItem.objCard
    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then
        Check_Square_Playment_IsValied = True
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True,objCard.接口序号=" & objCard.接口序号 & ",objCard.消费卡=" & objCard.消费卡)
        Exit Function
    End If
    
    On Error GoTo errHandle
    
 
    
    dblMoney = objCurItem.结算金额
    If dblMoney = 0 Then
        Check_Square_Playment_IsValied = True
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true,dblMoney=" & dblMoney)
        Exit Function
    End If
    If dblMoney < 0 Then
        Screen.MousePointer = 0
        MsgBox "收款金额未输入或不能输入负数,请检查!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,dblMoney=" & dblMoney)
        Exit Function
         
    End If
    
    dbl未付金额 = RoundEx(objBalanceInfo.未付合计 - objBalanceInfo.冲预交合计, 6)
    
    If dblMoney = 0 And dbl未付金额 <> 0 Then
        Screen.MousePointer = 0
        MsgBox "收款金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,dblMoney=" & dblMoney & ",dbl未付金额=" & dbl未付金额)
        Exit Function
    End If
    If dblMoney > Format(dbl未付金额, "0.00") And dblMoney <> 0 Then
        Screen.MousePointer = 0
        MsgBox "收款金额不能大于本次未付金额:" & Format(dbl未付金额, "0.00") & " ！", vbInformation, gstrSysName
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,dblMoney=" & dblMoney & ",dbl未付金额=" & Format(dbl未付金额, "0.00"))
        Exit Function
    End If
    
    '先检查对应的接口
    If mbytMode = 2 Then
        If mEditType = g_Ed_门诊结帐 Then
            If GetUninClassMoney(0, rsMoney) = False Then
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,GetUninClassMoney=false,mEditType=" & mEditType)
                Exit Function
            End If
        Else
            If GetUninClassMoney(objBalanceInfo.结帐ID, rsMoney) = False Then
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,GetUninClassMoney=false,objBalanceInfo.结帐ID=" & objBalanceInfo.结帐ID)
                Exit Function
            End If
        End If
    Else
        '先检查对应的接口
        If mEditType = g_Ed_门诊结帐 Or mEditType = g_Ed_住院结帐 Then
            If zlGetClassMoney(0, objBalanceInfo.当前结帐, rsMoney, mrsFeeList) = False Then
                 Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,GetUninClassMoney=false,mEditType=" & mEditType)
                Exit Function
            End If
        Else
            If zlGetClassMoney(objBalanceInfo.结帐ID, objBalanceInfo.当前结帐, rsMoney, mrsFeeList) = False Then
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,GetUninClassMoney=false,objBalanceInfo.结帐ID=" & objBalanceInfo.结帐ID)
                Exit Function
            End If
        End If
    End If
    Call zlGetBalanceItemsFromCardObject(vsBlance, objCard, objBalanceItemsOld)


   'mstr费用性质-本次支付的费用类型,多个用逗号分隔(如:1,2,3,4): 1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
    str费用性质 = Replace(mstr费用性质, "3", "1")
    str费用性质 = Replace(str费用性质, "4", "3")
    If mobjThirdSwap.zlSquare_Payment_IsValied(objPati, objBalanceInfo, objCurItem, objBalanceItems_out, dbl帐户余额, 0, objBalanceItemsOld, str费用性质, rsMoney) = False Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,mobjThirdSwap.zlSquare_Payment_IsValied=false")
        Exit Function
    End If
    
    Call ClearSquareBalance(objCard.接口序号)
    Call zlReCalcBalanceInfor(vsBlance, objBalanceInfo)
    
    Check_Square_Playment_IsValied = True
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true")
    Exit Function
errHandle:
    Screen.MousePointer = 0
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false,Err=" & Err.Description)
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Excute_Third_Payment(ByVal objPati As clsPatientInfo, ByRef objCurItems As clsBalanceItems, ByRef objBalanceInfor As clsBalanceInfo, _
    Optional ByRef blnSave_Out As Boolean, Optional objSetFocus As Object, Optional blnChangeMoney_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行三方扣款接口
    '入参:objCurItems-当前扣款对象
    '出参:blnSave_Out-是否保存了数据
    '     objSetFocus-返回移动焦点对象
    '     blnChangeMoney_Out-是否金额发生了变化
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-13 14:30:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCheck As Boolean, intStatus As Integer, dblOldMoney As Double, strErrMsg As String
    Dim objItem As clsBalanceItem, objItems As clsBalanceItems
    Dim cllPro As Collection, cllDeposit As Collection
    Dim i As Long, objCard As Card
    Dim objCurItem As clsBalanceItem
    On Error GoTo errHandle
    
    blnSave_Out = False
    blnChangeMoney_Out = False
    If objCurItems Is Nothing Then Excute_Third_Payment = True: Exit Function
    If objCurItems.Count = 0 Then Exit Function
    
    Set objItem = objCurItems(1)
    Set objCard = objItem.objCard
    If objItem.消费卡 Or objItem.卡类别ID = 0 Then Excute_Third_Payment = True: Exit Function
    
    dblOldMoney = objCurItems.结算金额
    
 
    '先检查是否合法
    blnCheck = True
    If objItem.是否保存 Then
        If Not mobjThirdSwap.zlThird_IsSwapIsSucces(objCurItems, intStatus, strErrMsg) Then
           If intStatus = 1 Then
                If strErrMsg <> "" Then
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行扣款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                Else
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行扣款处理中，请稍后再试" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                End If
                Exit Function
           End If
        Else
            blnCheck = False
        End If
    End If
    
    If blnCheck Then
        '1.三方卡检查
        Set objItems = Nothing
        Set objCurItem = objCurItems(1)
        If objCurItems.Count > 1 Then Set objCurItem = zlCopyNewItemFromBalanceItem(objCurItem): objCurItem.结算金额 = objCurItems.结算金额
        
        If Check_ThirdPayment_IsValied(objCurItem, objPati, objBalanceInfor, objItems) = False Then Exit Function
        '当前结算金额为0直接返回true(接口已经检查了未付不等于零的情况)
        If objCurItems.结算金额 = 0 Then Excute_Third_Payment = True: Exit Function
    End If
    
    If Not objItems Is Nothing Then
        Set objCurItems = objItems
        For i = 1 To objCurItems.Count
            objCurItems(i).是否保存 = objItem.是否保存
            objCurItems(i).是否结算 = objItem.是否结算
            objCurItems(i).是否强制退现 = objItem.是否强制退现
            objCurItems(i).是否退款 = objItem.是否退款
            objCurItems(i).是否退款分交易 = objItem.是否退款分交易
            objCurItems(i).是否预交 = objItem.是否预交
            objCurItems(i).是否允许编辑 = objItem.是否允许编辑
            objCurItems(i).是否允许删除 = objItem.是否允许删除
            objCurItems(i).是否允许退现 = objItem.是否允许退现
            objCurItems(i).是否转帐 = objItem.是否转帐
            objCurItems(i).退款交易流水号 = objItem.退款交易流水号
            objCurItems(i).退款交易说明 = objItem.退款交易说明
            objCurItems(i).限制类别 = objItem.限制类别
            objCurItems(i).消费卡 = objItem.消费卡
            objCurItems(i).消费卡ID = objItem.消费卡ID
            objCurItems(i).校对标志 = objItem.校对标志
            objCurItems(i).预交ID = objItem.预交ID
            objCurItems(i).找补 = objItem.找补
            objCurItems(i).门诊结帐 = objItem.门诊结帐
            objCurItems(i).卡类别ID = objItem.卡类别ID
            objCurItems(i).结帐时间 = objItem.结帐时间
            objCurItems(i).结帐ID = objItem.结帐ID
            objCurItems(i).结算类型 = objItem.结算类型
            objCurItems(i).结算IDs = objItem.结算IDs
            objCurItems(i).关联交易ID = objItem.关联交易ID
            objCurItems(i).缴款金额 = objItem.缴款金额
            objCurItems(i).行号 = objItem.行号
        Next
    End If
    
    '2.先数据处理
    Set cllPro = New Collection: Set cllDeposit = New Collection
     
    If GetSaveBalanceSQL(objPati, objBalanceInfor, cllPro) = False Then Exit Function
    
    If objBalanceInfor.是否保存预交 = False Then
        '预交款处理，需要先处理预交数据，已经处理的预交，不用再处理
        If GetSaveStrickDepositSQL(objBalanceInfor, cllDeposit) = False Then Exit Function
        For i = 1 To cllDeposit.Count
            cllPro.Add cllDeposit(i)
        Next
    End If
    
    
    If mobjThirdSwap.zlThird_Payment(objCard, objPati, cllPro, objBalanceInfor, objCurItems, objItems, blnSave_Out) = False Then
        If blnSave_Out Then objBalanceInfor.是否保存结帐单 = True
        If cllDeposit.Count <> 0 And blnSave_Out Then objBalanceInfor.是否保存预交 = True
        
        If blnSave_Out Then
            Call zlAddBalanceDataToGridFromBalanceItems(vsBlance, objCard, objBalanceInfor, objItems, objItem.行号)
        End If
        Call zlReCalcBalanceInfor(vsBlance, objBalanceInfor)
        RaiseEvent AfterPayMoney(objItems, blnSave_Out, False, False)
        Call zlSetBalanceRowDataFromItemsObject(vsBlance, objItems)       '设置行状态
        Exit Function
    End If
    
    
    If blnSave_Out Then objBalanceInfor.是否保存结帐单 = True
    If cllDeposit.Count <> 0 And blnSave_Out Then objBalanceInfor.是否保存预交 = True
    
    objItems.结算金额 = RoundEx(objItems.结算金额, 5)
    dblOldMoney = RoundEx(dblOldMoney, 5)
        
    Excute_Third_Payment = True
    If Not objItems Is Nothing Then
        '增加到支付列表中
        Call zlAddBalanceDataToGridFromBalanceItems(vsBlance, objCard, objBalanceInfor, objItems, objItem.行号)
        If RoundEx(dblOldMoney, 5) <> RoundEx(objItems.结算金额, 5) Then
            MsgBox "注意:" & vbCrLf & _
                    "    " & objCard.名称 & " 结算交易与你输入的金额不一致，请注意收款! " & vbCrLf & _
                    "输入结算金额:" & Format(dblOldMoney, "0.00") & vbCrLf & _
                    "当前交易金额:" & Format(objItems.结算金额, "0.00"), vbOKOnly + vbInformation, gstrSysName
            blnChangeMoney_Out = True
        
        End If
        vsBlance.Row = vsBlance.Rows - 1
    End If
    
    RaiseEvent AfterPayMoney(objItems, blnSave_Out, True, blnChangeMoney_Out)
    Call zlReCalcBalanceInfor(vsBlance, objBalanceInfor)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function GetUninClassMoney(ByRef lng结帐ID As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, dblMoney As Double
    Dim dblTemp As Double, cllFilter As Collection
    Dim str病人IDs As String, strSubTable As String, varPara() As Variant
    
    On Error GoTo errHandle
    
    '初始化数据结构
    Set rsMoney = New ADODB.Recordset
    rsMoney.Fields.Append "收费类别", adVarChar, 10, adFldIsNullable
    rsMoney.Fields.Append "金额", adDouble, , adFldIsNullable
    rsMoney.CursorLocation = adUseClient
    rsMoney.LockType = adLockOptimistic
    rsMoney.CursorType = adOpenStatic
    rsMoney.Open
        
    If lng结帐ID <> 0 Then
        strSql = "" & _
        "   Select  A.收费类别,nvl(sum(A.结帐金额) ,0) as 金额   " & _
        "   From 门诊费用记录 A" & _
        "   Where A.结帐ID=[1] Group by A.收费类别 " & _
        "   Union ALL " & _
        "   Select  A.收费类别,nvl(sum(A.结帐金额) ,0) as 金额   " & _
        "   From 住院费用记录 A" & _
        "   Where A.结帐ID=[1] Group by A.收费类别 "
        strSql = "Select 收费类别,Sum(金额) as 金额 From (" & strSql & ")  Group by  收费类别"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, mfrmMain.Caption, lng结帐ID)
    
        With rsTemp
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                rsMoney.Find "收费类别='" & Nvl(!收费类别, "无") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!收费类别 = Nvl(!收费类别, "无")
                rsMoney!金额 = Val(Nvl(rsMoney!金额)) + Val(Nvl(!金额))
                rsMoney.Update
                .MoveNext
            Loop
        End With
        GetUninClassMoney = True
        Exit Function
    End If
    
    Set cllFilter = New Collection
    cllFilter.Add Array("合同单位ID", mlng合约单位ID)
    If gobjSquare.objOneCardComLib.zlGetPatiIdsByRange("", str病人IDs, , , True, cllFilter) = False Then Exit Function
    If zlGetVarBoundSQL(0, str病人IDs, strSubTable, varPara, 0) = False Then Exit Function
    
    strSql = "Select A.收费类别, Nvl(Sum(实收金额), 0) - Nvl(Sum(结帐金额), 0) As 未结金额" & _
            " From 门诊费用记录 A, (" & strSubTable & ") B" & _
            " Where A.病人id = B.Column_Value And A.记录状态 <> 0 And A.记帐费用 = 1 And A.门诊标志 IN(1,4)" & _
            " Group By a.收费类别" & _
            " Having Nvl(Sum(a.实收金额), 0) - Nvl(Sum(a.结帐金额), 0) <> 0  "
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSql, mfrmMain.Caption, varPara)
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            rsMoney.Find "收费类别='" & Nvl(!收费类别, "无") & "'", , adSearchForward, 1
            If rsMoney.EOF Then rsMoney.AddNew
            rsMoney!收费类别 = Nvl(!收费类别, "无")
            rsMoney!金额 = Val(Nvl(rsMoney!金额)) + Val(Nvl(!未结金额))
            rsMoney.Update
            .MoveNext
        Loop
    End With
    GetUninClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ClearSquareBalance(ByVal lngCardTypeID As Long, Optional ByVal lng消费卡ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除消费卡结算
    '编制:刘兴洪
    '日期:2015-01-23 14:54:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, j As Long
    With vsBlance
        j = 1
        Do While j <= .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("类型"))) = 5 _
                And Val(.TextMatrix(j, .ColIndex("卡类别ID"))) = lngCardTypeID _
                And (lng消费卡ID = 0 Or (lng消费卡ID <> 0 And Val(.TextMatrix(j, .ColIndex("消费卡ID"))) = lng消费卡ID)) Then
                If .Rows >= 2 Then
                    .RowData(j) = ""
                    .RemoveItem j
                Else
                    .Rows = 2
                   .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                   .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                   .RowData(1) = ""
                   j = 2
                End If
            Else
                j = j + 1
            End If
        Loop
    End With
End Sub
Private Function GetSaveBalanceSQL(ByVal objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本次结帐的结帐相关的Sql
    '入参:objPati-病人信息
    '     objBalanceInfor-当前结帐数据集
    '出参:cllBillPro-保存结帐数据的SQL集
    '     objBalanceInfor-更新后的结帐数据集
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-13 11:10:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean
    RaiseEvent GetBalanceSaveSQL(objPati, objBalanceInfor, cllBillPro, blnCancel)
    If blnCancel = True Then Exit Function
    GetSaveBalanceSQL = True
End Function

Private Function GetSaveStrickDepositSQL(ByRef objBalanceInfor As clsBalanceInfo, ByRef cllDeposit As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保存预交款的数据
    '出参:cllDeposit-相关的数据集
    '     objSetFocus-获取失败时,缺省光标定位到指定的控件中
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-19 15:12:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean
    On Error GoTo errHandle
    RaiseEvent GetSaveStrickDepositSQL(objPatiInfor, objBalanceInfor, cllDeposit, blnCancel)
    If blnCancel = True Then Exit Function
    GetSaveStrickDepositSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub InitBalanceGrid(Optional bln查看 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算列表
    '编制:刘兴洪
    '日期:2015-01-23 14:14:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBlance
    
        For i = 1 To .Rows - 1
            .RowData(i) = ""
        Next
        .Clear: .Rows = 2: i = 0: .Cols = 20
        .TextMatrix(0, i) = "卡类别ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "消费卡ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算性质": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "编辑状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "类型": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否退现": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否全退": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "校对标志": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否密文": .ColWidth(i) = 0: i = i + 1
        
        .TextMatrix(0, i) = "结算方式": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "结算金额": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "结算号码": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "卡类别名称": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "卡号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易流水号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易说明": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "备注": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "关联交易ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否转账": .ColWidth(i) = 0: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            Select Case .ColKey(i)
            Case "是否转账", "关联交易ID", "结算性质", "类型", "是否保存", "是否密文", "校对标志", "编辑状态", "是否退现", "是否全退", "结算状态", "是否验证"
                .ColHidden(i) = True
                .ColData(i) = """-1||1"
            Case "结算金额"
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = """1||0"
            Case .ColIndex("结算方式")
                .ColData(i) = """1||0"
            Case "卡类别名称"
                .ColData(i) = "1||2"
            Case .ColIndex("结算号码")
                .ColData(i) = "1||0"
            Case Else
                .ColData(i) = "1||" & IIf(bln查看, "0", "2")
                
            End Select
            If bln查看 Then .ColData(i) = ""
        Next
        If Not bln查看 Then .Editable = flexEDKbdMouse
    End With
    zl_vsGrid_Para_Restore mlngModule, vsBlance, mfrmMain.Name, "结算列表"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


