VERSION 5.00
Begin VB.Form frmReplenishTheBalancePrint 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "frmReplenishTheBalancePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------------------------------------
'程序入口相关变量
Private mbytInFun As Byte                 '1-新单打印,2-重打,3-退费打印; 4-补打票据;6-退费票据(红票)打印
Private mobjFactProperty As clsFactProperty
Private mintInsure As Integer
Private mstrReclaimInvoice As String    '要求回收的发票号,按1-根据系统预定规则分配票号和2-根据用户自助规则分配票号有效
'--------------------------------------------------------------------------------------
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mlng领用ID As Long              '上次领用ID
Private mstrPrintNO As String           '要打印的单据号，多个时用逗号分隔:'F0000001','F0000002',...
Private mstrInvoice As String           '开始票据号
Private mdatFeeDate As Date             '费用单据数据的登记时间
Private mblnPrinted As Boolean          '票据数据生成是否成功(是否已打印)
Private mstrPrivs As String
Private mstrUseType As String
Private mbln分配票号 As Boolean
Private mobjInvoice As clsInvoice
Private mbln只用一张票据 As Boolean
Private mlngModule As Long

Private Type Ty_PrintSheet
    blnCalcMoney As Boolean '是否累计发票金额
    lngPrePage As Long '上一页页号
    lngGridCount As Long '当前页已打印表格个数
    lngCurPrintRow As Long '当前打印行数，不区分页数
    dblInvoiceMoney As Double '当前页累计发票金额
    arrInvoice As Variant '发票号，与页号一一对应
    blnUseOnlyOneInvoice As Boolean '是否仅使用一张发票
End Type
Private mPrintSheet As Ty_PrintSheet
 

Public Sub ReportPrint(ByVal bytInFun As Byte, ByVal strNos As String, ByVal intInsure As Integer, _
                        ByVal objFactProperty As clsFactProperty, _
                        ByVal strReclaimInvoice As String, _
                        ByRef lngLastUseID As Long, ByVal strInvoice As String, ByVal datFeeDate As Date, _
                        Optional blnVirtualPrint As Boolean, _
                        Optional ByVal blnDelRecord As Boolean, _
                        Optional blnPrintBillEmpty As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据打印,程序入口
    '入参:bytInfun :1-新单打印,2-重打,3-退费打印,4-补打票据(只有:2-按系统预定规则和3-用户自定规则时才转入),6-退费票据(红票)打印
    '       strNOs - 新单时要打印的单据号，多个时用逗号分隔:'F0000001','F0000002',...,
    '                   - 退费票据(红票)打印时，传入结算序号
    '       strReclaimInvoice-要求回收的发票号,多个用逗号分离'F0000001','F0000002',...
    '       lngLastUseID-最近使用的领用批次ID,初次时为0
    '       strInvoice-开始票据号，不带引号,不严格控制票据时允许传入空,严格控制时调用前已前检查不能为空
    '       datFeeDate-费用结算时间
    '       blnVirtualPrint-医保接口内调用打印，HIS只走票号不实际打印
    '       blnDelRecord-重打时，是否是对退费记录进行重打(目前只有北京医保(医保接口打印票据)才允许)
    '       lngShareUseID-共享批次
    '       strUseType-使用类别
    ' 出参:
    '   blnPrintBillEmpty-是否打印的空表数据()
    '编制:刘兴洪
    '日期:2014-09-24 18:11:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, strPrintNO As String, blnPrint As Boolean, blnTrans As Boolean
    Dim strReportNO As String, strSQL As String, strClearNOs As String, strFormat As String, lngBalanceID As Long
    Dim blnNotPrint As Boolean, varTmp As Variant '空变量，主要是为了处理调用函数的返回值
    Dim str发票号 As String, int票据张数 As Integer
    blnPrintBillEmpty = False
    mbln分配票号 = False
    
    mbytInFun = bytInFun: mdatFeeDate = datFeeDate: mlngModule = 1124
    
    
    mlng领用ID = lngLastUseID: mstrInvoice = strInvoice: mstrReclaimInvoice = strReclaimInvoice
    Set mobjFactProperty = objFactProperty: mintInsure = intInsure
    If bytInFun <> 6 Then strNos = IIf(InStr(1, strNos, "'") = 0, "'" & Replace(strNos, ",", "','") & "'", strNos)
    
    Me.Caption = "打印"
    
    '1.变量传递
    If mbytInFun = 6 Then '退费票据(红票)打印
        strReportNO = "ZL" & glngSys \ 100 & "_BILL_1124_3"
    Else
        strReportNO = "ZL" & glngSys \ 100 & "_BILL_1124"
    End If
    strFormat = IIf(objFactProperty.打印格式 = 0, "", "ReportFormat=" & objFactProperty.打印格式)
    
    mstrPrintNO = "": mblnPrinted = False
    blnNotPrint = (Not gobjTax Is Nothing And gblnTax) Or blnVirtualPrint
    
    '2.打印调用
    Select Case mbytInFun
        Case 1 '新单打印或重打票据
             
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp) '调用打印方法但不打印，只生成了票据使用数据
                If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice '修改时,只清除新单据的开始票据号
                Call TaxInterface(1, mstrPrintNO, "")
            Else
               '票据接口
                If BillPrint(1, mstrPrintNO, "", "", strClearNOs) = False Then: GoTo ClearInvoice
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "发票号=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty
                If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice
            End If
 
        Case 2, 4 '重打
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
                Call TaxInterface(2, mstrPrintNO, "")       '打印税控票据
                ''调用医保重打接口
                If InsureReprint(blnVirtualPrint, Replace(Split(strNos, ",")(0), "'", ""), lngBalanceID, blnDelRecord, strInvoice) = False Then Exit Sub
            Else
                '票据接口
                 If BillPrint(2, mstrPrintNO, "", strInvoice, strClearNOs) = False Then Exit Sub
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "发票号=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
        Case 3  '退费
        
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
                Call TaxInterface(3, mstrPrintNO, "")
            Else
                If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit Sub
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "发票号=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", "", "", strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
        Case 6 '红票打印
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
'                Call TaxInterface(3, mstrPrintNO, "")
            Else
'                If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit Sub
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "结算序号=" & Val(strNos), "PrintEmpty=0", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
    End Select
    
    '3.传回最近使用的领用ID
    lngLastUseID = mlng领用ID
    Exit Sub
ClearInvoice:
    On Error GoTo errH
    
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(Split(strClearNOs, ","))
            strPrintNO = Split(strClearNOs, ",")(i)
            strSQL = "Zl_票据起始号_Update('" & strPrintNO & "','',1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mlngModule = 1124
    mstrPrivs = ";" & GetPrivFunc(glngSys, mlngModule)
    Set mobjReport = New clsReport
    Set mobjInvoice = New clsInvoice
    Call mobjInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    
    With mPrintSheet
        If .blnCalcMoney = False Then Exit Sub
        
        If .lngPrePage > 0 Then
            If .blnUseOnlyOneInvoice Then
                Call UpdateInvoiceMoney(.arrInvoice(0), .dblInvoiceMoney)
            Else
                '保存最后一页的数据
                Call UpdateInvoiceMoney(.arrInvoice(.lngPrePage - 1), .dblInvoiceMoney)
            End If
        End If
    End With
End Sub

Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrInvoice As Variant)
    Dim strSQL As String, i As Integer, strInvoices As String
    
    With mPrintSheet
        .blnCalcMoney = True
        .lngPrePage = 0
        .lngCurPrintRow = 0
        .blnUseOnlyOneInvoice = False
    End With
    
    If mbln只用一张票据 Then
        mPrintSheet.blnUseOnlyOneInvoice = True
        TotalPages = 1 '收费每次打印只用一张票据
    End If
    
    '没有票据号,严格控制票据时不打印,不严格控制票据时只打印不处理票据数据
    If mstrInvoice = "" Then
        Cancel = mobjFactProperty.严格控制
        mblnPrinted = Not mobjFactProperty.严格控制
        mPrintSheet.blnCalcMoney = False '不计算票据金额
        Exit Sub
    End If
    
    
    If CheckInvoiceValied(TotalPages, mbytInFun = 6) = False Then Cancel = True: Exit Sub
    
    On Error GoTo errH
    
    '2.生成票据数据
    Select Case mbytInFun
        Case 1
            strSQL = "Zl_补充结算票据_Insert('" & Replace(mstrPrintNO, "'", "") & "','" & mstrInvoice & "'," & ZVal(mlng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                     "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),0," & TotalPages & ")"
        Case 2, 3
            '如果是多张，只需要传一张单据号就行了(修改多张中的一张时,最后一张是新的)
            strSQL = "Zl_补充结算票据_Reprint('" & Replace(Split(mstrPrintNO, ",")(0), "'", "") & "','" & mstrInvoice & "'," & ZVal(mlng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                    "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & IIf(mbytInFun = 2, "0", "1") & "," & TotalPages & ")"
        Case 6 '退费发票(红票)
            'Zl_补充结算退费票据_Insert
            strSQL = "Zl_补充结算退费票据_Insert("
            '  结算序号_In   病人预交记录.结算序号%Type,
            strSQL = strSQL & "" & Val(mstrPrintNO) & ","
            '  票据号_In       票据使用明细.号码%Type,
            strSQL = strSQL & "'" & mstrInvoice & "',"
            '  领用id_In       票据使用明细.领用id%Type,
            strSQL = strSQL & "" & ZVal(mlng领用ID) & ","
            '  使用人_In       票据使用明细.使用人%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  使用时间_In     票据使用明细.使用时间%Type,
            strSQL = strSQL & "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  票据张数_In Number:=1
            strSQL = strSQL & "" & TotalPages & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSQL, "票据数据生成")
    mblnPrinted = True
    
    '3.传递所用的票据号信息
    For i = 1 To TotalPages
        strInvoices = strInvoices & "," & mstrInvoice
        If i < TotalPages Then mstrInvoice = zlStr.Increase(mstrInvoice)
    Next
    strInvoices = Mid(strInvoices, 2)
    arrInvoice = Split(strInvoices, ",")
    
    mPrintSheet.arrInvoice = arrInvoice
        
    '不严格控制票据时保存到注册表
    If Not mobjFactProperty.严格控制 Then
        zlDatabase.SetPara "当前收费票据号", mstrInvoice, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub
Private Function CheckInvoiceValied(Optional int张数 As Integer = 1, _
    Optional ByVal blnDelFeePrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查发票是否合法(严格控制票据时)
    '入参:int张数 -需要的发票张数
    '   blnDelFeePrint-退费发票(红票)打印
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-24 17:49:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjFactProperty.严格控制 Then CheckInvoiceValied = True: Exit Function
    
    '1.严格控制票据时，根据实际的票据张数,重新检查领用ID和票据号Property.使用类别, mlng领用
    If mobjInvoice.zlGetInvoiceGroupID(mlngModule, UserInfo.姓名, EM_收费收据, mobjFactProperty.使用类别, mobjFactProperty.共享批次ID, mlng领用ID, mlng领用ID, int张数, mstrInvoice) = False Then Exit Function
    '数据合法
    If mlng领用ID > 0 Then CheckInvoiceValied = True: Exit Function
    Select Case mlng领用ID
        Case -1
            MsgBox IIf(blnDelFeePrint, "本次退费发票(红票)打印", "单据[" & mstrPrintNO & "]") & "共需要" & int张数 & "张票据！" & vbCrLf & _
                "你没有足够的自用和共用的票据，请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
        Case -2
            MsgBox IIf(blnDelFeePrint, "本次退费发票(红票)打印", "单据[" & mstrPrintNO & "]") & "共需要" & int张数 & "张票据！" & vbCrLf & _
                "你没有足够的的共用票据，请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
        Case -3
            MsgBox IIf(blnDelFeePrint, "本次退费发票(红票)打印", "单据[" & mstrPrintNO & "]") & "共需要" & int张数 & "张票据！" & vbCrLf & _
                "票据号[" & mstrInvoice & "]不在可用领用批次的有效票据号范围内！" & _
                "请重新输入有效的票据号后重打该单据！", vbInformation, gstrSysName
        Case -4
            MsgBox IIf(blnDelFeePrint, "本次退费发票(红票)打印", "单据[" & mstrPrintNO & "]") & "共需要" & int张数 & "张票据！" & vbCrLf & _
                "票据号[" & mstrInvoice & "]所在的领用批次没有足够的票据！" & _
                "请先打印其它票据,用完当前领用批次后，重打该单据！", vbInformation, gstrSysName
        Case Else
            MsgBox "票据领用信息访问失败！将来，你可以" & IIf(blnDelFeePrint, "重打该单据！", "重打单据[" & mstrPrintNO & "]！"), vbInformation, gstrSysName
    End Select
End Function


Private Sub TaxInterface(ByVal byt类型 As Byte, ByVal strPrintNO As String, ByVal strModiNos As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用税控打印接口
    '入参:byt类型-1-正常打印(含修改);2-重打;3-退费
    '        strPrintNO-要打印的单据号，多个时用逗号分隔:'F0000001','F0000002',...
    '        strModiNos-修改多单据中的一张时,指该多张单据的所有NO，用逗号分隔:'F0000001','F0000002',...
    '编制:刘兴洪
    '日期:2013-03-27 14:24:03
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    '未启用税控,直接返回
    If Not gblnTax Then Exit Sub
    If byt类型 = 3 Then
        '退费
        gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strPrintNO)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        gstrTax = gobjTax.zlTaxOutReput(gcnOracle, strPrintNO)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If byt类型 = 2 Then
        '重打
        MsgBox "请在准备好之后按确定开始打印。", vbInformation, gstrSysName
        gstrTax = gobjTax.zlTaxOutReput(gcnOracle, strPrintNO)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If strModiNos <> "" Then
        gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strModiNos)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
    End If
    gstrTax = gobjTax.zlTaxOutPrint(gcnOracle, strPrintNO)
    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub
Private Function BillPrint(ByVal byt类型 As Byte, ByVal strPrintNO As String, _
    ByVal strModiNos As String, ByRef strInvoice As String, ByRef strClearNOs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用票据打印接口
    '入参:byt类型-1-正常打印(或修改打印);2-重打打印;3-退费
    '        strPrintNO-要打印的单据号，多个时用逗号分隔:'F0000001','F0000002',...
    '        strModiNos-修改多单据中的一张时,指该多张单据的所有NO，用逗号分隔:'F0000001','F0000002',...
    '         strInvoice-发票号(重打时有效)
    '出参:strClearNOs-需要清除的单据号
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-03-27 14:36:28
    '问题:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gblnBillPrint Then BillPrint = True: Exit Function
    If byt类型 = 3 Then
        '退费
        '退费事务之前先调了票据收回：zlEraseBill
        BillPrint = gobjBillPrint.zlRePrintBill(strPrintNO, 0, strInvoice)
        Exit Function
    End If
    If byt类型 = 2 Then
        '重打
       BillPrint = gobjBillPrint.zlRePrintBill(strPrintNO, 0, strInvoice)
       Exit Function
    End If
    If strModiNos <> "" Then
        If gobjBillPrint.zlEraseBill(strModiNos, 0) = False Then strClearNOs = Replace(strModiNos, "'", ""): Exit Function
    End If
    If gobjBillPrint.zlPrintBill(strPrintNO, 0) = False Then strClearNOs = Replace(strPrintNO, "'", ""): Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InsureReprint(ByVal blnVirtualPrint As Boolean, ByVal strNos As String, _
    ByVal lng结帐ID As Long, ByVal bln退费 As Boolean, ByRef strInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新调用医保打印接口
    '入参:blnVirtualPrint-是否调用医保接口打印
    '       strNos-单据号
    '       bln退费-是否退费
    '       strInvoice-发票号
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-24 18:02:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    If Not blnVirtualPrint Then InsureReprint = True: Exit Function
    '81222
    If lng结帐ID = 0 Then
        strSQL = "Select Max(结算ID) As 结算ID From 费用补充记录 Where NO= [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
        If Not rsTmp.EOF Then
            lng结帐ID = rsTmp!结算ID
        End If
    End If
    Call gclsInsure.RePrintBill(mintInsure, lng结帐ID, strInvoice)
    InsureReprint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mobjReport_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long)
    
    On Error GoTo errHandle
    '约定[0列0宽度]为票据金额
    If Sheet Is Nothing Then Exit Sub
    If Sheet.COLS = 0 Then Exit Sub
    If Sheet.ColWidth(0) <> 0 Then Exit Sub
    
    With mPrintSheet
        If .blnCalcMoney = False Then Exit Sub
        
        If .lngPrePage <> Page Then
            If .lngPrePage > 0 And .blnUseOnlyOneInvoice = False Then
                '当前页号变化了且不是打印值使用一张发票，则保存上一页的数据
                Call UpdateInvoiceMoney(.arrInvoice(.lngPrePage - 1), .dblInvoiceMoney)
                .dblInvoiceMoney = 0
            ElseIf .lngPrePage = 0 Then
                .dblInvoiceMoney = 0
            End If
            
            .lngPrePage = Page
            .lngGridCount = 0
        End If
        
        '含有多个表格时，以第一个表格为准
        If Row = 1 Then .lngGridCount = .lngGridCount + 1
        If .lngGridCount > 1 Then Exit Sub
        
        '累计金额
        .dblInvoiceMoney = .dblInvoiceMoney + Val(Sheet.TextMatrix(.lngCurPrintRow, 0))
        
        '累计表格行号，表格对象是不区分页数的
        .lngCurPrintRow = .lngCurPrintRow + 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub UpdateInvoiceMoney(ByVal strInvoice As String, ByVal dblMoney As Double)
    '更新票据金额
    Dim strSQL As String
    
    On Error GoTo errHandle
    'Zl_票据使用明细_更新金额
    strSQL = "Zl_票据使用明细_更新金额("
    '  领用id_In   票据使用明细.领用id%Type,
    strSQL = strSQL & "" & mlng领用ID & ","
    '  发票号_In   票据使用明细.号码%Type,
    strSQL = strSQL & "'" & strInvoice & "',"
    '  票据金额_In 票据使用明细.票据金额%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  票种_In     票据使用明细.票种%Type := 1
    strSQL = strSQL & "" & 1 & ")"
    zlDatabase.ExecuteProcedure strSQL, "更新票据金额"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
