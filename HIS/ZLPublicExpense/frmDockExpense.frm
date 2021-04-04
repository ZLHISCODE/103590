VERSION 5.00
Begin VB.Form frmDockExpense 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "医嘱附费管理"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   Icon            =   "frmDockExpense.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin zlPublicExpense.ctlDockExpense dkeExpense 
      Height          =   4860
      Left            =   240
      TabIndex        =   0
      Top             =   60
      Width           =   6660
      _ExtentX        =   12330
      _ExtentY        =   8625
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
Attribute VB_Name = "frmDockExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'相关事件
Public Event Activate() '自已激活时
Public Event RequestRefresh() '要求主窗体刷新
Event StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String) '要求更新主窗体状态栏文字
'bytType:1-附费执行,2-附费取消
Public Event zlPopupMenu(lng医嘱ID As Long, lng发送号 As Long, strNO As String, int记录性质 As Integer, X As Single, Y As Single)
'------------------------------------------


Private mfrmParent As Object
Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤消出院 = 17            '允许撤消病人出院
    support必须录入入出诊断 = 18    '病人入院与出院时，必须录入诊断名
    support记帐完成后上传 = 19      '要求上传在记帐数据提交后再进行
    support出院结算必须出院 = 20    '病人结帐时如果选择出院结帐，就检查必须出院才可以进行
    
    support挂号使用个人帐户 = 21    '使用医保挂号时是否使用个人帐户进行支付

    support门诊连续收费 = 22        '门诊在身份验证后，可进行多次收费操作
    support门诊收费完成后验证 = 23  '在门诊收费完成，是否再次调用身份验证
    
    support医嘱上传 = 24            '医嘱产生费用时是否实时传输
    support分币处理 = 25            '医保病人是否处理分币
    support中途结算仅处理已上传部分 = 26 '提供对已上传部分数据的结算功能
    support允许冲销已结帐的记帐单据 = 27 '是否允许冲销记帐单据，如果该单据已经结帐
    
    support允许部份冲销单据 = 28
    support出院无实际交易 = 29 '出院接口中是否要与接口商进行交易
    support负数记帐 = 35            '是否允许负数记帐，操作员首先要拥有负数记帐的权限。此参数缺省为真，不支持的接口需单独处理
    support医保接口打印票据 = 46    'HIS中只走票据号但不调打印，医保接口(北京)中打印
    support医生确定处方类型 = 48
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    support实时监控 = 60
    
    support上传门诊档案 = 70                    '在门诊医嘱发送时，是否调用TranElecDossier函数完成门诊病人电子卷宗/电子档案的上传
    support挂号不收取病历费 = 81    '在挂号时，不使用医保收取病历费
    support挂号检查项目 = 86
End Enum
Private mbytFontSize As Byte
Private mlngOptModule As Long '附费模块号
Private mlngPlugInID As Long '自动执行的插件功能ID
Private mrsPlugInBar As ADODB.Recordset '菜单样式
Private mstrPreAdviceIdAndPayNums As String
Private mobjSaveData As Object
Private mblnFirst As Boolean
Private mstrMainPrivs As String
Private mcbsMain As Object
Private mobjSquareCard As Object

Public Sub SetFontSize(ByVal bytSize As Byte)
      '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘兴洪
    '日期:2012-06-18 16:50:35
    '问题:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    '对于vsFlexGrid控件在使用个性化设置时会加大列宽，因此在窗体初次加载是不设置字体,主要是getForm方法引起
    
    dkeExpense.FontSize = mbytFontSize
End Sub
 
Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object, ByRef objSquareCard As Object)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    Set mfrmParent = frmParent
    If cbsMain Is Nothing Then Exit Sub
    
    If Not mblnFirst Then
        mblnFirst = True
        If objSquareCard Is Nothing Then
            Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
            If mobjSquareCard.zlInitComponents(Me, p医嘱附费管理, glngSys, gstrDBUser, gcnOracle, False) = False Then
                Set mobjSquareCard = Nothing
                MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
            End If
        Else
            Set mobjSquareCard = objSquareCard
        End If
        Set mcbsMain = cbsMain
        Set cbsMain.Icons = gobjCommFun.GetPubIcons
        Call GetPlugInBar(p医嘱附费管理, -1, mrsPlugInBar)
    End If
    
    '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "费用(&M)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "生成主费用(&N)") '不计价时显示为:补充主费用,手工计价时显示为:生成主费用
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_NewItem, "补充附加费用(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改费用(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除费用(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ExtraFeeMove, "附费转移(&T)")
        objControl.IconId = conMenu_Edit_CollectMan
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ExtraFeeExe, "附费执行(&E)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ExtraFeeUnExe, "附费取消执行(&F)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "销帐申请(&L)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelAudit, "销帐审核(&U)")
        '外挂菜单
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End With
    
    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "医嘱附费选项(&O)"): objControl.BeginGroup = True
        objControl.IconId = conMenu_File_Parameter
    End With

    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    Set objBar = cbsMain(2)
    For Each objControl In objBar.Controls '先求出前面的最后一个Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    With objBar.Controls
        'Set objControl = .Find(, conMenu_File_Preview) '从预览按钮之后开始加入
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "生成主费", objControl.Index + 1): objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_NewItem, "补附费", objControl.Index + 1): objPopup.BeginGroup = True
            objPopup.ID = conMenu_Edit_NewItem: objPopup.IconId = conMenu_Edit_NewItem
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "改费", objPopup.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删费", objControl.Index + 1)
                
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ExtraFeeExe, "附费执行", objControl.Index + 1): objControl.BeginGroup = True
        objControl.IconId = conMenu_Edit_Leave_Modify
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ExtraFeeUnExe, "取消执行", objControl.Index + 1)
        objControl.IconId = conMenu_Edit_Leave_Delete
    End With
    
    '命令的快键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyE, conMenu_Edit_Append '生成主费用
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改附加费用
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete '删除附加费用
    End With

    '设置不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
    End With
    
    
    '外挂程序对象初始化
    Call DefCommandPlugIn(cbsMain, mrsPlugInBar)
End Sub

Private Sub DefCommandPlugIn(ByRef cbsMain As Object, ByRef rsBar As ADODB.Recordset)
'功能：外挂部件菜单接入。
'说明：判断关键字     InTool 决定菜单样式
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim i As Long
    Dim lngTmp As Long
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    '独立按钮
    rsBar.Filter = "IsInTool=1  and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        If Not objMenu Is Nothing Then
            With objMenu.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                        objControl.IconId = rsBar!图标ID
                        objControl.Parameter = rsBar!功能名
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '下拉按钮，如果只有一个按钮，也当作独立按钮
    rsBar.Filter = "IsInTool=0 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        If Not objMenu Is Nothing Then
            Set objPopup = objMenu.CommandBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "扩展功能", , False)
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                    objControl.IconId = rsBar!图标ID
                    objControl.Parameter = rsBar!功能名
                    objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '工具栏按钮
    Set objBar = cbsMain(2)
    Set objControl = objBar.FindControl(, conMenu_Help_Help)
    If Not objControl Is Nothing Then
        objControl.BeginGroup = True
        lngTmp = objControl.Index - 1
    Else
        lngTmp = -1
    End If
    rsBar.Filter = "IsInTool=1 and BarType=2"
    If Not rsBar.EOF Then
        With objBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!功能名, lngTmp + 1)
                    objControl.IconId = rsBar!图标ID
                    objControl.Parameter = rsBar!功能名
                    objControl.Style = xtpButtonIconAndCaption
                lngTmp = objControl.Index
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                rsBar.MoveNext
            Next
            objControl.BeginGroup = True
        End With
    End If
    rsBar.Filter = "IsInTool=0 and BarType=2"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        Set objPopup = objBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "扩展功能", lngTmp + 1, False)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.IconId = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        lngTmp = objPopup.Index
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名, lngTmp + 1)
                objControl.IconId = rsBar!图标ID
                objControl.Parameter = rsBar!功能名
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                lngTmp = objPopup.Index
                rsBar.MoveNext
            Next
        End With
    End If
    '自动执行的功能
    rsBar.Filter = "IsAuto=1"
    If Not rsBar.EOF Then mlngPlugInID = rsBar!功能ID
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim strNO As String
    
    Select Case Control.ID
    Case conMenu_File_PrintSet '打印设置
        Call zlPrintSet
    Case conMenu_File_Preview '预览费用清单
        Call dkeExpense.zlPrintData(2)
    Case conMenu_File_Print '打印费用清单
        Call dkeExpense.zlPrintData(1)
    Case conMenu_File_Excel '输出费用清单
        Call dkeExpense.zlPrintData(3)
        
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Edit_Append '生成主费用
        Call dkeExpense.zlBuildMainExpense(Me)
    Case conMenu_Edit_NewItem * 10# + 1 '补费:收费单据
         Call dkeExpense.zlFuncFeeNewPrice(Me)
    Case conMenu_Edit_NewItem * 10# + 2 '补费:记帐单据
        Call dkeExpense.zlFuncFeeNewBilling(Me)
    Case conMenu_Edit_NewItem * 10# + 3 '补费:零耗费用
        Call dkeExpense.zlFuncFeeNewNull(Me)
    Case conMenu_Edit_NewItem * 10# + 4  '补费:备份卫材划价收费
       Call dkeExpense.zlFuncStuffCharge(Me, 1)
    Case conMenu_Edit_NewItem * 10# + 5 ''备份卫材记帐
       Call dkeExpense.zlFuncStuffCharge(Me, 2)
    Case conMenu_Edit_Modify '修改附费
        Call dkeExpense.zlFuncFeeModi(Me)
    Case conMenu_Edit_Delete '删除附费
        Call dkeExpense.zlFuncFeeDel(Me)
    Case conMenu_Edit_ExtraFeeMove '附费转移
        Call dkeExpense.zlFuncExtraFeeMove(Me)
    Case conMenu_Edit_ExtraFeeExe   '附费执行
        Call dkeExpense.zlFuncExtraFeeExe(Me, 1, mstrMainPrivs)
    Case conMenu_Edit_ExtraFeeUnExe '附费取消执行
        Call dkeExpense.zlFuncExtraFeeExe(Me, 0, mstrMainPrivs)
    Case conMenu_Edit_ChargeDelApply
        Call dkeExpense.zlFuncAdviceReCharge(1, Me)
    Case conMenu_Edit_ChargeDelAudit   '销帐申请审核
        Call dkeExpense.zlFuncAdviceReCharge(2, Me)
    Case conMenu_Tool_Option '医嘱附费选项
        If frmExpenseSetup.zlEditCard(mfrmParent) Then
            '刷新费用明细
            dkeExpense.Refresh
        End If
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '外挂功能执行
        Call dkeExpense.zlFuncPlugIn(Me, Control)
    Case Else
    End Select
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Call dkeExpense.zlUpdateCommandBars(mcbsMain, Control)
End Sub
Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    Dim objControl As CommandBarControl

    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_NewItem '补费
        With CommandBar.Controls
            .DeleteAll
            '扩1位,为了使用快捷键
            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 1, "收费单据(&1)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 2, "记帐单据(&2)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 3, "零耗费用(&3)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 4, "备货卫材收费(&3)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 5, "备货卫材记帐(&4)")
            With mcbsMain.KeyBindings
                .Add FCONTROL, vbKeyN, conMenu_Edit_NewItem * 10# + 1
                .Add FCONTROL, vbKeyB, conMenu_Edit_NewItem * 10# + 2
                .Add FCONTROL, vbKeyS, conMenu_Edit_NewItem * 10# + 4
            End With
        End With
    End Select
End Sub
Private Sub dkeExpense_Activate()
    RaiseEvent Activate
End Sub
Private Sub dkeExpense_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub
 
Private Sub dkeExpense_StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String)
    RaiseEvent StatusTextUpdate(bytType, Text)
End Sub

Private Sub dkeExpense_zlPopupMenu(lng医嘱ID As Long, lng发送号 As Long, strNO As String, int记录性质 As Integer, X As Single, Y As Single)
    RaiseEvent zlPopupMenu(lng医嘱ID, lng发送号, strNO, int记录性质, X, Y)
End Sub

Private Sub Form_Load()
    mbytFontSize = 9
    mblnFirst = False
    Set mrsPlugInBar = Nothing
    Call dkeExpense.zlInitCommon(mobjSquareCard)
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With dkeExpense
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub

Public Function zlRefresh(ByVal lng科室id As Long, ByVal strAdviceIdAndPayNums As String, _
    Optional ByVal blnMoved As Boolean = False, Optional ByVal strNos As String, _
    Optional ByVal byt记录性质 As Byte, Optional ByVal byt病人来源 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据重新刷新
    '入参:lng科室id-科室ID
    '     strAdviceIdAndPayNums-医嘱ID和发送号和独立执行标志(医嘱ID1:发送号1:独立执行,医嘱ID2:发送号2:独立执行,...)
    '     strNos:单据号(多个传入时,用逗号分离)
    '     byt记录性质:医嘱ID传空时,才传入,单据性质(1-收费单;2-记帐单)
    '     byt病人来源-1-门诊;2-住院
    '     blnMoved -该病人的数据是否已转出
    '     bln单独执行-用于检验项目，一并采集的一组项目，是否针对其中的某一个单独执行
    '出参:
    '编制:刘兴洪
    '日期:2014-04-10 11:02:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objControl As CommandBarControl
    
    zlRefresh = dkeExpense.zlRefresh(Me, lng科室id, strAdviceIdAndPayNums, blnMoved, strNos, byt记录性质, byt病人来源)
    
    If mstrPreAdviceIdAndPayNums <> strAdviceIdAndPayNums And mlngPlugInID <> 0 Then
        Set objControl = mcbsMain.FindControl(, mlngPlugInID, , True)
        If Not objControl Is Nothing Then objControl.Execute
        mstrPreAdviceIdAndPayNums = strAdviceIdAndPayNums
    End If
End Function

Public Function zlBuildMainExpense() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:生成主费用
    '返回:生成成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBuildMainExpense = dkeExpense.zlBuildMainExpense
End Function

Public Function zlAddChargeExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt应用场合 As Byte, _
    Optional ByVal lng病人ID As Long, _
    Optional ByVal lng开单科室id As Long, Optional ByVal lng病人科室ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补收费费用
    '入参:frmMain-调用主窗体
    '     lngModule-模块号
    '     byt应用场合:0-医嘱附费;1-体检补费(可选参数)
    '出参:strOutNos-成功保存的单据号
    '返回:补收费费用,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.医嘱附费执行具体的功能(参见:zlCisKernel.dockExpense)
    '       不需要传入病人ID; 开单科室及病人科室ID
    '    2.体检补费时,需要传入lng病人ID,开单科室ID,病人科室ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If byt应用场合 = 0 Then
        zlAddChargeExpense = dkeExpense.zlFuncFeeNewPrice(frmMain, strOutNos, objSaveData)
        Exit Function
    End If
    If frmTechnicExpense.EditCard(frmMain, GetInsidePrivs(p医嘱附费管理), 0, 0, 0, lng病人ID, 0, _
         1, 1, lng开单科室id, lng病人科室ID, 0, "", "", "", "", , , False, strOutNos, byt应用场合, objSaveData, mobjSquareCard) Then
         zlAddChargeExpense = True
    End If
End Function

Public Function zlAddBillingExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt应用场合 As Byte, ByVal int病人来源 As Integer, _
    Optional ByVal lng病人ID As Long, Optional lng主页Id As Long, _
    Optional ByVal lng开单科室id As Long, _
    Optional ByVal lng病人科室ID As Long, Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补记帐费用
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    byt应用场合:0-医嘱附费;1-体检补费(可选参数)
    '    int病人来源:1-门诊病人,2-住院病人
    '出参:strOutNos-成功保存的单据号
    '返回:补记帐费用,补费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.医嘱附费执行具体的功能(参见:zlCisKernel.dockExpense)
    '       不需要传入病人ID;开单科室及病人科室ID
    '    2.体检补费时,需要传入lng病人ID,开单科室ID,病人科室ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If byt应用场合 = 0 Then
        '99053:李南春,2016/7/29，调用补记账费用接口
        zlAddBillingExpense = dkeExpense.zlFuncFeeNewBilling(frmMain, strOutNos)
        Exit Function
    End If
    zlAddBillingExpense = frmTechnicExpense.EditCard(frmMain, GetInsidePrivs(p医嘱附费管理), 0, 0, 0, lng病人ID, lng主页Id, _
          int病人来源, 2, lng开单科室id, lng病人科室ID, 0, "", "", "", "", , , False, strOutNos, byt应用场合, objSaveData, mobjSquareCard)
End Function

Public Function zlAddZeroExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt应用场合 As Byte, ByVal int病人来源 As Integer, _
    Optional ByVal lng病人ID As Long, Optional ByVal lng主页Id As Long, _
    Optional ByVal lng开单科室id As Long, _
    Optional ByVal lng病人科室ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补零费用
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    byt应用场合:0-医嘱附费;1-体检补费(可选参数)
    '    int病人来源:1-门诊病人,2-住院病人
    '    lng主页ID -医嘱附费和门诊病人传入0,住院病人必须转入
    '出参:strOutNos-返回零耗费单号,多个用逗号分离
    '返回:补记帐费用,补费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.医嘱附费执行具体的功能(参见:zlCisKernel.dockExpense)
    '       不需要传入病人ID;开单科室及病人科室ID
    '    2.体检补费时,需要传入lng病人ID,开单科室ID,病人科室ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If byt应用场合 = 0 Then
        zlAddZeroExpense = dkeExpense.zlFuncFeeNewNull(frmMain, strOutNos, objSaveData)
        Exit Function
    End If
    zlAddZeroExpense = frmTechnicExpense.EditCard(frmMain, GetInsidePrivs(p医嘱附费管理), 0, 0, 0, lng病人ID, lng主页Id, _
          int病人来源, 2, lng开单科室id, lng病人科室ID, 0, "", "", "", "", , , True, strOutNos, byt应用场合, objSaveData, mobjSquareCard)
End Function

Public Function zlAddStuffChargeExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt应用场合 As Byte, ByVal int病人来源 As Integer, _
    Optional ByVal lng病人ID As Long, _
    Optional ByVal lng开单科室id As Long, _
    Optional ByVal lng病人科室ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补备货卫材收费费用
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    byt应用场合:0-医嘱附费;1-体检补费(可选参数)
    '    int病人来源:1-门诊病人,2-住院病人
    '出参:strNo-返回备货卫材单号,多个用逗号分离
    '返回:补记帐费用,补费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.医嘱附费执行具体的功能(参见:zlCisKernel.dockExpense)
    '       不需要传入病人ID;开单科室及病人科室ID
    '    2.体检补费时,需要传入lng病人ID,开单科室ID,病人科室ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If byt应用场合 = 0 Then
        Set mobjSaveData = objSaveData
        zlAddStuffChargeExpense = dkeExpense.zlFuncStuffCharge(frmMain, 1, strOutNos, objSaveData)
        Exit Function
    End If
    zlAddStuffChargeExpense = frmStuffCharge.zlBillEdit(frmMain, 0, lngModule, GetInsidePrivs(p医嘱附费管理), 1, "", _
         1, lng病人ID, 0, lng开单科室id, lng病人科室ID, _
          0, "", False, "", 0, 0, "", , , strOutNos, byt应用场合, objSaveData, mobjSquareCard) = True
End Function

Public Function zlAddStuffBillingExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt应用场合 As Byte, ByVal int病人来源 As Integer, _
    Optional ByVal lng病人ID As Long, _
    Optional ByVal lng主页Id As Long, _
    Optional ByVal lng开单科室id As Long, _
    Optional ByVal lng病人科室ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补备货卫材记帐费用
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    byt应用场合:0-医嘱附费;1-体检补费(可选参数)
    '    int病人来源:1-门诊病人,2-住院病人
    '    lng主页ID-主页ID可以不传(但住院病人一定要传入)
    '出参:strNo-返回备货卫材单号,多个用逗号分离
    '返回:补记帐费用,补费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.医嘱附费执行具体的功能(参见:zlCisKernel.dockExpense)
    '       不需要传入病人ID;开单科室及病人科室ID
    '    2.体检补费时,需要传入lng病人ID,开单科室ID,病人科室ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If byt应用场合 = 0 Then
        Set mobjSaveData = objSaveData
        zlAddStuffBillingExpense = dkeExpense.zlFuncStuffCharge(frmMain, 2, strOutNos, objSaveData)
        Exit Function
    End If
    zlAddStuffBillingExpense = frmStuffCharge.zlBillEdit(frmMain, 0, lngModule, GetInsidePrivs(p医嘱附费管理), 1, "", _
         2, lng病人ID, 0, lng开单科室id, lng病人科室ID, _
          0, "", False, "", 0, 0, "", , , strOutNos, byt应用场合, objSaveData, mobjSquareCard) = True
End Function
Public Function zlIsFunValied(ByVal bytType As Byte, ByVal bytPrivsCheck As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判检查某功能是否有效
    '入参: bytType- 1-修改附费;2-删除附费;3-附费转移;4-附费执行;5-附费取消执行;6-销帐申请;7-销帐审核
    '      bytPrivsCheck -检查权限:0-不检查权限;1-检查数据和权限;2-仅检查权限
    '出参:
    '返回:功能有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 17:00:52
    '说明:
    '   1.根据附费列表中的内容,检查某项功能是否有效
    '   2.根据权限检查某项功能是否有效
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlIsFunValied = dkeExpense.IsFunValied(bytType, bytPrivsCheck)
End Function

Public Function zlModifyExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal int病人来源 As Integer, ByVal int记录性质 As Integer, ByVal strNO As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改附费
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    int病人来源-病人来源
    '    int记录性质
    '    strNO
    '返回:修改附费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.执行具体的功能(参见:zlCisKernel.dockExpense),不需要传入记录性质和NO
    '    2.体检调用时,需要传入记录性质和NO
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlModifyExpense = dkeExpense.zlFuncFeeModi(frmMain, int病人来源, int记录性质, strNO, objSaveData)
End Function
Public Function zlDelExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, Optional int病人来源 As Integer, _
    Optional ByVal int记录性质 As Integer, _
    Optional ByVal strNO As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除附费
    '入参:frmMain -调用主窗体
    '    int病人来源-1-门诊;2-住院
    '    lngModule -模块号
    '    int记录性质
    '    strNO
    '返回:修改附费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.执行具体的功能(参见:zlCisKernel.dockExpense),不需要传入记录性质和NO
    '    2.体检调用时,需要传入,病人来源,记录性质和NO
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlDelExpense = dkeExpense.zlFuncFeeDel(frmMain, int病人来源, int记录性质, strNO, objSaveData)
End Function
Public Function zlMoveExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:附费移动
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '返回:附费移动成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlMoveExpense = dkeExpense.zlFuncExtraFeeMove(frmMain)
End Function
Public Function zlExcuteExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, Optional bln取消执行 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:附费执行
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '返回:附费执行成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'bytType=0-取消执行,1-执行
    zlExcuteExpense = dkeExpense.zlFuncExtraFeeExe(frmMain, IIf(bln取消执行, 0, 1), mstrMainPrivs)
End Function


Public Function zlApplyExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:销帐申请
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '返回:销帐申请成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlApplyExpense = dkeExpense.zlFuncAdviceReCharge(1, frmMain)
End Function
 
Public Function zlAuditExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:销帐申请审核
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '返回:销帐申请审核成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlAuditExpense = dkeExpense.zlFuncAdviceReCharge(2, frmMain)
End Function
Public Function zlParaOptionSet(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:销帐申请审核
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '返回:销帐申请审核成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
   zlParaOptionSet = frmExpenseSetup.zlEditCard(frmMain)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mobjSaveData Is Nothing Then Set mobjSaveData = Nothing
End Sub

Private Sub DefCommandPlugInPopup(ByVal objBar As Object, ByRef rsBar As ADODB.Recordset)
'功能：在医嘱卡右键弹出菜单
    Dim i As Long
    Dim objControl As CommandBarControl
    Dim objCtl As CommandBarControl
    Dim objPopup As CommandBarPopup
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    '独立按钮
    rsBar.Filter = "IsInTool=1 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        For i = 1 To rsBar.RecordCount
            Set objControl = objBar.Add(xtpControlButton, rsBar!功能ID, rsBar!功能名)
            objControl.IconId = rsBar!图标ID
            objControl.Parameter = rsBar!功能名
            objControl.Style = xtpButtonIconAndCaption
            If Val(rsBar!IsGroup) = 1 Then
                objControl.BeginGroup = True
            End If
            rsBar.MoveNext
        Next
    End If
    
    rsBar.Filter = "IsInTool=0 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        Set objPopup = objBar.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "扩展功能")
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                objControl.IconId = rsBar!图标ID
                objControl.Parameter = rsBar!功能名
                objControl.Style = xtpButtonIconAndCaption
                If Val(rsBar!IsGroup) = 1 Then
                    objControl.BeginGroup = True
                End If
                rsBar.MoveNext
            Next
        End With
    End If
End Sub

Private Function GetPlugInBar(ByVal lng模块 As Long, ByVal int场合 As Integer, rsBar As ADODB.Recordset) As String
'功能：组织外挂部件的菜单样按钮
    Dim strFunc As String
    Dim strXML As String
    Call CreatePlugIn(lng模块, int场合)
    If gobjPlugIn Is Nothing Then Exit Function
    On Error Resume Next
    strFunc = gobjPlugIn.GetFuncNames(glngSys, lng模块, int场合, strXML)
    Call zlPlugInErrH(Err, "GetFuncNames")
    Err.Clear: On Error GoTo 0
    Call MakePlugInBar(strFunc, strXML, rsBar)
    GetPlugInBar = strFunc
End Function

Private Sub MakePlugInBar(ByVal strFunc As String, ByVal strXML As String, rsBar As ADODB.Recordset)
'功能：组织菜单到本地记录集中，注意对老版本的兼容处理
'参数：strFunc 老版本功能列串，strXML含配置信息的功能串
    Dim strM As String
    Dim strB As String
    Dim strP As String
    Dim strTag As String
    Dim i As Long
    Dim strTmp As String
    Dim lngS As Long, lngE As Long
    Dim rsBarFuncID As ADODB.Recordset
    
    If strXML = "" And strFunc <> "" Then
        '兼容以前老版本的方式
        Call InitPlugInRsBar(rsBar)
        Call AddPlugInBarRs(rsBar, strFunc, 1)
        Call AddPlugInBarRs(rsBar, strFunc, 2)
        Call AddPlugInBarRs(rsBar, strFunc, 3)
        Call SetPlugInBar(rsBar, 1)
        Exit Sub
    End If
    
    On Error GoTo errH
    strXML = Trim(strXML)
    '暂定为200个扩展功能插件，防止死循环
    For i = 0 To 200
        lngS = InStr(strXML, "<")
        lngE = InStr(strXML, ">")
        strTag = Mid(strXML, lngS + 1, lngE - lngS - 1)
        If strTag = "menubar" Then
            lngS = lngE
            lngE = InStr(strXML, "</menubar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strM = strM & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "toolbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</toolbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strB = strB & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "popbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</popbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strP = strP & "," & strTmp
            strXML = Mid(strXML, lngE + 9)
        End If
        If strXML = "" Then
            Exit For
        End If
    Next
    If strM = "" Then Exit Sub
    strM = Mid(strM, 2)
    strB = Mid(strB, 2)
    strP = Mid(strP, 2)

    Call InitPlugInRsBar(rsBar)
    Call AddPlugInBarRs(rsBar, strM, 1)
    Call AddPlugInBarRs(rsBar, strB, 2)
    Call AddPlugInBarRs(rsBar, strP, 3)
    Call SetPlugInBar(rsBar, 2)
    Exit Sub
errH:
    If 2 = 1 Then
        Resume
    End If
End Sub

Private Sub AddPlugInBarRs(ByRef rsBar As ADODB.Recordset, ByVal strFunc As String, ByVal intType As Integer)
'功能：将功能串转换为记录集方式
'参数：strFunc 功能串，intType 功能按钮属于那一栏 1-菜单栏，2-工具栏，3-左键栏
    Dim varFunc As Variant
    Dim i As Long
    Dim strFuncName As String
    Dim blnFirstTool As Boolean
    If strFunc = "" Then Exit Sub
    varFunc = Split(strFunc, ",")
    With rsBar
        For i = 0 To UBound(varFunc)
            strFuncName = varFunc(i)
            .AddNew
            !BarType = intType
            If InStr(strFuncName, "Auto:") > 0 Then
                !IsAuto = 1
                strFuncName = Replace(strFuncName, "Auto:", "")
            Else
                !IsAuto = 0
            End If
            
            If InStr(strFuncName, "InTool:") > 0 Then
                !IsInTool = 1
                strFuncName = Replace(strFuncName, "InTool:", "")
            Else
                !IsInTool = 0
            End If
            If InStr(strFuncName, "|:") > 0 Then
                !IsGroup = 1
                strFuncName = Replace(strFuncName, "|:", "")
            Else
                !IsGroup = 0
                If Not blnFirstTool And !IsInTool = 1 Then
                    '第一个独立按钮显示分割线
                    blnFirstTool = True
                    !IsGroup = 1
                End If
            End If
            !功能名 = strFuncName
            !菜单名 = strFuncName
            .Update
        Next
    End With
End Sub

Private Function SetPlugInBar(ByRef rsBar As ADODB.Recordset, ByVal lngV As Long) As String
'功能：分配功能ID，加菜单快键
'参数：lngV 版本，1-老版，2-新版
'返回：字符串，以前低版本方式的功能串
    Dim i As Long
    '分配功能ID，图标ID
    With rsBar
        .Filter = 0
        If .EOF Then Exit Function
        .MoveFirst
        For i = 1 To .RecordCount
            !序号 = i
            !功能ID = conMenu_Tool_PlugIn_Item + i
            !图标ID = conMenu_Tool_PlugIn_Item
            If lngV = 1 Then
                !IsInTool = 0
                !IsGroup = 0
            End If
            .Update
            .MoveNext
        Next
    End With
    Call SetPlugInBarKey(rsBar, 1, lngV)
    Call SetPlugInBarKey(rsBar, 2, lngV)
    Call SetPlugInBarKey(rsBar, 3, lngV)
    rsBar.Filter = 0
End Function

Private Sub SetPlugInBarKey(rsBar As ADODB.Recordset, ByVal intType As Integer, ByVal lngV As Long)
'功能：设定快键
'参数：lngV 版本，1-老版，2-新版 intType 功能按钮属于那一栏 1-菜单栏，2-工具栏，3-左键栏
    Dim i As Long
    With rsBar
        .Filter = "IsInTool=0 and BarType=" & intType
        If .RecordCount = 1 And lngV = 2 Then
            '如果只有一个，也归为独立按钮
            !IsInTool = 1
            .Update
        Else
            For i = 1 To .RecordCount
                If i <= 35 Then
                    If i <= 9 Then
                        !菜单名 = !菜单名 & "(&" & i & ")"
                    Else
                        !菜单名 = !菜单名 & "(&" & Chr(55 + i) & ")"
                    End If
                    .Update
                    .MoveNext
                Else
                    Exit For
                End If
            Next
        End If
        
        .Filter = "IsInTool=1 and BarType=" & intType
        For i = 1 To .RecordCount
            If i <= 35 Then
                If i <= 9 Then
                    !菜单名 = !菜单名 & "(&" & i & ")"
                Else
                    !菜单名 = !菜单名 & "(&" & Chr(55 + i) & ")"
                End If
                .Update
                .MoveNext
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub InitPlugInRsBar(rsBar As ADODB.Recordset)
    Set rsBar = New ADODB.Recordset
    rsBar.Fields.Append "序号", adBigInt '用于排序
    rsBar.Fields.Append "功能ID", adBigInt '菜单按钮 Control.ID
    rsBar.Fields.Append "图标ID", adBigInt
    rsBar.Fields.Append "功能名", adVarChar, 1000 '去掉关键字之后的 名称 即工具栏上的按钮名称
    rsBar.Fields.Append "菜单名", adVarChar, 1000 '菜单栏/右键菜单 名称
    rsBar.Fields.Append "IsAuto", adInteger '是否自动执行功能
    rsBar.Fields.Append "IsGroup", adInteger '是否分割线
    rsBar.Fields.Append "IsInTool", adInteger '是否独立显示
    rsBar.Fields.Append "BarType", adInteger '1-菜单栏，2－工具栏，3－弹出栏
    rsBar.CursorLocation = adUseClient
    rsBar.LockType = adLockOptimistic
    rsBar.CursorType = adOpenStatic
    rsBar.Open
End Sub
