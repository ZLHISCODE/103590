VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "票据打印"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mbytInFun As Byte                 '1-新单打印,2-重打,3-退费打印; 4-补打票据;6-退费票据(红票)打印
Private mlng领用ID As Long              '上次领用ID
Private mstrPrintNO As String           '要打印的单据号，多个时用逗号分隔:'F0000001','F0000002',...
Private mstrInvoice As String           '开始票据号
Private mdatFeeDate As Date             '费用单据数据的登记时间
Private mblnPrinted As Boolean          '票据数据生成是否成功(是否已打印)
Private mstrReclaimInvoice As String    '要求回收的发票号,按1-根据系统预定规则分配票号和2-根据用户自助规则分配票号有效
Private mlngShareUseID As Long '打印的共享领用ID
Private mstrUseType As String
Private mbln分配票号 As Boolean

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

Private Sub Form_Load()
    Set mobjReport = New clsReport
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
    
    '没有票据号,严格控制票据时不打印,不严格控制票据时只打印不处理票据数据
    If mstrInvoice = "" Then
        Cancel = gblnBill挂号
        mblnPrinted = Not gblnBill挂号
        mPrintSheet.blnCalcMoney = False '不计算票据金额
        Exit Sub
    End If
    
    If CheckInvoiceValied(TotalPages, mbytInFun = 6) = False Then Cancel = True: Exit Sub
    
    On Error GoTo errH
    '2.生成票据数据
    Select Case mbytInFun
        Case 1, 4
            strSQL = "Zl_病人挂号票据_Insert("
            '  No_In           Varchar2,
            strSQL = strSQL & "'" & Replace(mstrPrintNO, "'", "") & "'" & ","
            '  票据号_In       票据使用明细.号码%Type,
            strSQL = strSQL & "'" & mstrInvoice & "',"
            '  领用id_In       票据使用明细.领用id%Type,
            strSQL = strSQL & "" & ZVal(mlng领用ID) & ","
            '  使用人_In       票据使用明细.使用人%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  使用时间_In     票据使用明细.使用时间%Type,
            strSQL = strSQL & "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  票据张数_In     Number := 1,
            strSQL = strSQL & "" & TotalPages & ","
            '  医保接口打印_In Number := 0,
            strSQL = strSQL & "0,"
            '  收费票据_In Number:=0
            strSQL = strSQL & "" & IIf(gblnSharedInvoice, 1, 0) & ")"
        Case 2, 3
            '如果是多张，只需要传一张单据号就行了(修改多张中的一张时,最后一张是新的)
            strSQL = "Zl_病人挂号记录_Reprint('" & Replace(Split(mstrPrintNO, ",")(0), "'", "") & "','" & mstrInvoice & "'," & ZVal(mlng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                    "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & IIf(mbytInFun = 2, "1", "0") & _
                    "," & TotalPages & ",'" & mstrReclaimInvoice & "'," & IIf(gblnSharedInvoice, 1, 0) & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSQL, "票据数据生成")
    mblnPrinted = True
    
    '3.传递所用的票据号信息
    For i = 1 To TotalPages
        strInvoices = strInvoices & "," & mstrInvoice
        If i < TotalPages Then mstrInvoice = zlCommFun.IncStr(mstrInvoice)
    Next
    strInvoices = Mid(strInvoices, 2)
    arrInvoice = Split(strInvoices, ",")
    
    mPrintSheet.arrInvoice = arrInvoice
    
    strSQL = "Zl_凭条打印记录_Update(4,'" & mstrPrintNO & "',1,'" & UserInfo.姓名 & "','发票号:" & strInvoices & "')"
    zlDatabase.ExecuteProcedure strSQL, "凭条打印记录"

    '不严格控制票据时保存到注册表
    '更新本地票据
    If Not gblnBill挂号 Then
        If gblnSharedInvoice Then
            zlDatabase.SetPara "当前收费票据号", mstrInvoice, glngSys, 1121
        Else
            zlDatabase.SetPara "当前挂号票据号", mstrInvoice, glngSys, 1111
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub

Private Function CheckInvoiceValied(Optional int张数 As Integer = 1, _
    Optional ByVal blnDelFeePrint As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查发票是否合法(严格控制票据时)
    '入参:int张数 -需要的发票张数
    '     blnDelFeePrint-退费发票(红票)打印
    '出参:
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-03-27 13:01:41
    '问题:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gblnBill挂号 Then CheckInvoiceValied = True: Exit Function
    '1.严格控制票据时，根据实际的票据张数,重新检查领用ID和票据号
    mlng领用ID = GetInvoiceGroupID(IIf(gblnSharedInvoice, 1, 4), int张数, mlng领用ID, mlngShareUseID, mstrInvoice, mstrUseType)
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
    '日期:2013-03-27 17:01:02
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer
    On Error GoTo errHandle
    If Not blnVirtualPrint Then InsureReprint = True: Exit Function
    Call gclsInsure.RePrintBill(intInsure, lng结帐ID, strInvoice)
    InsureReprint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ReportPrint(ByVal bytInFun As Byte, ByVal strNos As String, ByVal strReclaimInvoice As String, _
                        ByRef lngLastUseID As Long, ByVal lngShareUseID As Long, ByVal strInvoice As String, _
                        ByVal datFeeDate As Date, _
                        Optional str缴款 As String, Optional str找补 As String, _
                        Optional intPrintFormat As Integer, Optional blnVirtualPrint As Boolean, _
                        Optional ByVal blnDelRecord As Boolean, Optional strUseType As String = "", _
                        Optional blnPrintBillEmpty As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据打印,程序入口
    '入参:bytInfun :1-新单打印,2-退费打印,3-重打,4-补打票据(只有:2-按系统预定规则和3-用户自定规则时才转入),6-退费票据(红票)打印
    '       strNOs - 新单时要打印的单据号，多个时用逗号分隔:'F0000001','F0000002',...,
    '                   - 修改时,传入新单据号,只有一张,用于打印取消后清除开始票据号
    '                   - 退费票据(红票)打印时，传入结算序号
    '       strReclaimInvoice-要求回收的发票号,多个用逗号分离'F0000001','F0000002',...
    '       lngLastUseID-最近使用的领用批次ID,初次时为0
    '       strInvoice-开始票据号，不带引号,不严格控制票据时允许传入空,严格控制时调用前已前检查不能为空
    '       datFeeDate-费用单据数据的登记时间
    '       intPrintFormat-打印格式(打印格式序号)
    '       blnVirtualPrint-医保接口内调用打印，HIS只走票号不实际打印
    '       blnDelRecord-重打时，是否是对退费记录进行重打(目前只有北京医保(医保接口打印票据)才允许)
    '       lngShareUseID-共享批次
    '       strUseType-使用类别
    '       lng打印ID-传入的打印ID(blnOnePatiPrint=true时传入),报表可以根据打印ID从“临时票据打印内容”的临时表中来获取对应的收费单据
    '                 之所有要临时表，主要原因是因为按病人打印时，单据号可能会造成，而自助定义报表有所限制
    ' 出参:
    '   blnPrintBillEmpty-是否打印的空表数据()
    '编制:刘兴洪
    '日期:2011-04-29 12:01:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
  
    Dim i As Integer, j As Integer, strPrintNO As String, blnPrint As Boolean, blnTrans As Boolean
    Dim strReportNO As String, strSQL As String, strClearNOs As String, strFormat As String, lngBalanceID As Long
    Dim blnNotPrint As Boolean, varTmp As Variant '空变量，主要是为了处理调用函数的返回值
    Dim str发票号 As String, int票据张数 As Integer
    
    Me.Caption = "打印" '触发Form_Load事件
    blnPrintBillEmpty = False
    mbln分配票号 = False
    '1.变量传递
    mlngShareUseID = lngShareUseID
    mbytInFun = bytInFun
    mlng领用ID = lngLastUseID: mstrUseType = strUseType
    mstrInvoice = strInvoice
    mdatFeeDate = datFeeDate
    mstrReclaimInvoice = strReclaimInvoice
    strReportNO = "ZL" & glngSys \ 100 & "_BILL_1111"
    strFormat = IIf(intPrintFormat = 0, "", "ReportFormat=" & intPrintFormat)
    mstrPrintNO = ""
    mblnPrinted = False
    blnNotPrint = blnVirtualPrint
    '2.打印调用
    Select Case mbytInFun
        Case 1 '新单打印
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp) '调用打印方法但不打印，只生成了票据使用数据
                If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice '修改时,只清除新单据的开始票据号
            Else
                If gblnBillPrint Then
                    On Error Resume Next
                    If gobjBillPrint.zlPrintBill_Reg("'" & strNos & "'") = False Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice
                    On Error GoTo errH
                End If
               '票据接口
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "发票号=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", str缴款, str找补, strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice
            End If
        Case 3, 4 '重打、补打
            mstrPrintNO = strNos
            
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
                ''调用医保重打接口
                If InsureReprint(blnVirtualPrint, Replace(Split(strNos, ",")(0), "'", ""), lngBalanceID, blnDelRecord, strInvoice) = False Then Exit Sub
            Else
                If gblnBillPrint Then
                    On Error Resume Next
                    If mbytInFun = 3 Then '重打调用票据作废
                        If gobjBillPrint.zlEraseBill_Reg("'" & strNos & "'") = False Then Exit Sub
                    End If
                    If gobjBillPrint.zlRePrintBill_Reg("'" & strNos & "'", "'" & strInvoice & "'") = False Then Exit Sub
                    On Error GoTo errH
                End If
                '票据接口
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "发票号=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
        Case 2  '退费
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
            Else
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "发票号=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", "", "", strFormat, 2)
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
            strSQL = "Zl_票据起始号_Update('" & strPrintNO & "','',4)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjReport_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long)
    
    On Error GoTo errHandle
    '约定[0列0宽度]为票据金额
    If Sheet Is Nothing Then Exit Sub
    If Sheet.Cols = 0 Then Exit Sub
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
    strSQL = strSQL & "" & IIf(gblnSharedInvoice, 1, 4) & ")"
    zlDatabase.ExecuteProcedure strSQL, "更新票据金额"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
