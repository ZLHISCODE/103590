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
Private mlngRestNum As Long '上次领用批次剩余的票据张数
Private mlngNextUseID As Long '下一个可用的票据领用ID
Private mstrNextInvoice As String '下一个可用的票据批次的开始票据号
Private mblnOnePatiPrint As Boolean, mlng打印ID As Long '按病人补打票据时使用

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
    
    If mbytInFun <> 6 Then
        '56963
        If gTy_Module_Para.byt票据分配规则 <> 0 And mbln分配票号 Then
            mPrintSheet.blnCalcMoney = False
            Exit Sub
        End If
        
        If gTy_Module_Para.bln一张票据 And mblnOnePatiPrint = False Then
            mPrintSheet.blnUseOnlyOneInvoice = True
            TotalPages = 1 '收费每次打印只用一张票据
        End If
    End If
    '没有票据号,严格控制票据时不打印,不严格控制票据时只打印不处理票据数据
    If mstrInvoice = "" Then
        Cancel = gblnStrictCtrl
        mblnPrinted = Not gblnStrictCtrl
        mPrintSheet.blnCalcMoney = False '不计算票据金额
        Exit Sub
    End If
    If gblnStrictCtrl Then
        If zlCheckInvoiceValied(mlng领用ID, TotalPages, mstrInvoice, mlngShareUseID, mstrUseType, _
                                mlngRestNum, mlngNextUseID, mstrNextInvoice) = False Then Cancel = True: Exit Sub
    End If
    
    On Error GoTo errH
    '2.生成票据数据
    Select Case mbytInFun
        Case 1
          'Create Or Replace Procedure Zl_门诊收费票据_Insert
            strSQL = "Zl_门诊收费票据_Insert("
            '  No_In           Varchar2,
            strSQL = strSQL & "" & IIf(mblnOnePatiPrint, "NULL", "'" & Replace(mstrPrintNO, "'", "") & "'") & ","
            '  票据号_In       票据使用明细.号码%Type,
            strSQL = strSQL & "'" & mstrInvoice & "',"
            '  领用id_In       票据使用明细.领用id%Type,
            strSQL = strSQL & "" & ZVal(mlng领用ID) & ","
            '  使用人_In       票据使用明细.使用人%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  使用时间_In     票据使用明细.使用时间%Type,
            strSQL = strSQL & "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  打印id_In       票据打印内容.Id%Type := 0,
            strSQL = strSQL & "" & IIf(mblnOnePatiPrint, mlng打印ID, 0) & ","
            '  票据张数_In     Number := 1,
            strSQL = strSQL & "" & TotalPages & ","
            '  Next领用id_In   票据使用明细.领用id%Type := 0,
            strSQL = strSQL & "" & mlngNextUseID & ","
            '  Next票据号_In   票据使用明细.号码%Type := Null,
            strSQL = strSQL & "'" & mstrNextInvoice & "',"
            '  医保接口打印_In Number := 0,
            strSQL = strSQL & "0,"
            '  按病人打印_In Number:=0
            strSQL = strSQL & "" & IIf(mblnOnePatiPrint, 1, 0) & ")"
        Case 2, 3
            '如果是多张，只需要传一张单据号就行了(修改多张中的一张时,最后一张是新的)
            strSQL = "zl_门诊收费记录_RePrint('" & Replace(Split(mstrPrintNO, ",")(0), "'", "") & "','" & mstrInvoice & "'," & ZVal(mlng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                    "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & IIf(mbytInFun = 2, "0", "1") & "," & TotalPages & "," & _
                    "Null,1," & mlngNextUseID & ",'" & mstrNextInvoice & "')"
        Case 6 '退费发票(红票)打印
            'Zl_门诊退费票据_Insert
            strSQL = "Zl_门诊退费票据_Insert("
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
        If i < TotalPages Then
            If i = mlngRestNum Then
                mstrInvoice = mstrNextInvoice
            Else
                mstrInvoice = zlStr.Increase(mstrInvoice)
            End If
        End If
    Next
    strInvoices = Mid(strInvoices, 2)
    arrInvoice = Split(strInvoices, ",")
    
    mPrintSheet.arrInvoice = arrInvoice
        
    '不严格控制票据时保存到注册表
    If Not gblnStrictCtrl Then
        zlDatabase.SetPara "当前收费票据号", mstrInvoice, glngSys, 1121
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub

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
    BillPrint = True
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
    '日期:2013-03-27 17:01:02
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer
    On Error GoTo errHandle
    If Not blnVirtualPrint Then InsureReprint = True: Exit Function
    intInsure = ChargeExistInsure(strNos, 0, lng结帐ID, , bln退费)
    Call gclsInsure.RePrintBill(intInsure, lng结帐ID, strInvoice)
    InsureReprint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Sub ReportPrint(ByVal bytInFun As Byte, ByVal strNos As String, ByVal strAllNOs As String, ByVal strReclaimInvoice As String, _
                        ByRef lngLastUseID As Long, ByVal lngShareUseID As Long, ByVal strInvoice As String, _
                        ByVal datFeeDate As Date, _
                        Optional str缴款 As String, Optional str找补 As String, Optional bln分别打印 As Boolean, _
                        Optional intPrintFormat As Integer, Optional blnVirtualPrint As Boolean, _
                        Optional ByVal blnDelRecord As Boolean, Optional strUseType As String = "", _
                        Optional blnPrintBillEmpty As Boolean, _
                        Optional blnOnePatiPrint As Boolean, Optional lng打印ID As Long, _
                        Optional strPriceGrade As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据打印,程序入口
    '入参:bytInfun :1-新单打印,2-重打,3-退费打印,4-补打票据(只有:2-按系统预定规则和3-用户自定规则时才转入),6-退费票据(红票)打印
    '       strNOs - 新单时要打印的单据号，多个时用逗号分隔:'F0000001','F0000002',...,
    '                   - 修改时,传入新单据号,只有一张,用于打印取消后清除开始票据号
    '                   - 退费票据(红票)打印时，传入结算序号
    '       strAllNOs-修改多单据中的一张时,指该多张单据的所有NO，用逗号分隔:'F0000001','F0000002',...
    '       strReclaimInvoice-要求回收的发票号,多个用逗号分离'F0000001','F0000002',...
    '       lngLastUseID-最近使用的领用批次ID,初次时为0
    '       strInvoice-开始票据号，不带引号,不严格控制票据时允许传入空,严格控制时调用前已前检查不能为空
    '       datFeeDate-费用单据数据的登记时间
    '       intPrintFormat-打印格式(打印格式序号)
    '       blnVirtualPrint-医保接口内调用打印，HIS只走票号不实际打印
    '       blnDelRecord-重打时，是否是对退费记录进行重打(目前只有北京医保(医保接口打印票据)才允许)
    '       lngShareUseID-共享批次
    '       strUseType-使用类别
    '       blnOnePatiPrint-按病人补打票据(不分结算次数)
    '       lng打印ID-传入的打印ID(blnOnePatiPrint=true时传入),报表可以根据打印ID从“临时票据打印内容”的临时表中来获取对应的收费单据
    '                 之所有要临时表，主要原因是因为按病人打印时，单据号可能会造成，而自助定义报表有所限制
    '       strPriceGrade-价格等级，用于计算工本费
    ' 出参:
    '   blnPrintBillEmpty-是否打印的空表数据()
    '编制:刘兴洪
    '日期:2011-04-29 12:01:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, strPrintNO As String, blnPrint As Boolean, blnTrans As Boolean
    Dim strReportNO As String, strSQL As String, strClearNOs As String, strFormat As String, lngBalanceID As Long
    Dim blnNotPrint As Boolean, varTmp As Variant '空变量，主要是为了处理调用函数的返回值
    Dim str发票号 As String, int票据张数 As Integer, var发票号 As Variant
    blnPrintBillEmpty = False
    mbln分配票号 = False
    '1.变量传递
    mlngShareUseID = lngShareUseID
    mbytInFun = bytInFun
    mlng领用ID = lngLastUseID: mstrUseType = strUseType
    mstrInvoice = strInvoice
    mdatFeeDate = datFeeDate
    mstrReclaimInvoice = strReclaimInvoice
    mblnOnePatiPrint = blnOnePatiPrint: mlng打印ID = lng打印ID

    If mbytInFun = 6 Then '退费票据(红票)打印
        strReportNO = "ZL" & glngSys \ 100 & "_BILL_1121_7"
    Else
        strReportNO = "ZL" & glngSys \ 100 & "_BILL_1121_1"
        If strPriceGrade = "" Then strPriceGrade = "-" '特殊处理
    End If
    strFormat = IIf(intPrintFormat = 0, "", "ReportFormat=" & intPrintFormat)
    
    mstrPrintNO = ""
    mblnPrinted = False
    blnNotPrint = (Not gobjTax Is Nothing And gblnTax) Or blnVirtualPrint
    '2.打印调用
    Select Case mbytInFun
        Case 1 '新单打印,修改重打或重打票据
            If gTy_Module_Para.byt票据分配规则 <> 0 Then
                '1.根据系统预定规则打印票据;2-根据用户自定义分配票号
                '先分配票号:
                If gblnStrictCtrl Then
                    '模拟计算，把需要的票据张数算出来
                    If zlExeCuteBillNoSplit(True, 1, mlng领用ID, strNos, 0, mstrInvoice, mdatFeeDate, 1, str发票号, int票据张数, , , mlng打印ID) = False Then
                        strClearNOs = Replace(strNos, "'", "")
                        GoTo ClearInvoice:
                        Exit Sub
                    End If
                    If zlCheckInvoiceValied(mlng领用ID, int票据张数, mstrInvoice, mlngShareUseID, mstrUseType, _
                                         mlngRestNum, mlngNextUseID, mstrNextInvoice) = False Then
                        strClearNOs = Replace(strNos, "'", "")
                        GoTo ClearInvoice: Exit Sub
                    End If
                End If

                If zlExeCuteBillNoSplit(False, 1, mlng领用ID, strNos, 0, mstrInvoice, mdatFeeDate, 1, str发票号, int票据张数, _
                                    mlngNextUseID, mstrNextInvoice, mlng打印ID) = False Then
                    strClearNOs = Replace(strNos, "'", "")
                    GoTo ClearInvoice:
                    Exit Sub
                End If
                If int票据张数 = 0 Then
                    strClearNOs = Replace(strNos, "'", "")
                    GoTo ClearInvoice:
                    Exit Sub     '没有生成票据,直接返回
                End If
                mstrReclaimInvoice = str发票号
                mbln分配票号 = True
               
                '不严格控制票据时保存当前收费票据号
                If Not gblnStrictCtrl And mstrReclaimInvoice <> "" Then
                    var发票号 = Split(mstrReclaimInvoice, ",")
                    zlDatabase.SetPara "当前收费票据号", var发票号(UBound(var发票号)), glngSys, 1121
                End If
                
                If blnNotPrint Then
                   Call TaxInterface(1, mstrPrintNO, strAllNOs)      '打印税控票据
                Else
                    '票据接口
                     If BillPrint(1, mstrPrintNO, strAllNOs, "", strClearNOs) = False Then Exit Sub
                     '调用报表
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                    "发票号=" & str发票号, "NO='NO'", "打印ID=" & mlng打印ID, "价格等级=" & strPriceGrade, "PrintEmpty=0", str缴款, str找补, strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                End If
                Exit Sub
            End If
            If bln分别打印 And UBound(Split(strNos, ",")) > 0 And strAllNOs = "" Then
            '如果是修改的多张中的一张，即使现在参数是分别打印，仍然要一起打（因为原始是一起打的）。
                For i = 0 To UBound(Split(strNos, ","))
                    mblnPrinted = False '须在这里初始，因为可能打印部件在调BeforePrint之前就出错返回了
                    mstrPrintNO = Split(strNos, ",")(i)
                    blnPrint = True
                    If gTy_Module_Para.bln工本费 Then
                        '一张单据只有工本费不打印
                        If BillOnlyFactMoney(Replace(mstrPrintNO, "'", "")) Then blnPrint = False
                    End If
                    If blnPrint Then
                        If blnNotPrint Then
                            Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)    '调用打印方法但不打印，只生成了票据使用数据
                            If Not mblnPrinted Then Exit For
                            Call TaxInterface(1, mstrPrintNO, "")        '打印税控票据
                            '票据接口
                            If BillPrint(1, mstrPrintNO, "", "", "") = False Then Exit For
                        Else
                            '票据接口
                            If BillPrint(1, mstrPrintNO, "", "", "") = False Then Exit For
                            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                                "发票号=FactNo", "NO=" & mstrPrintNO, "打印ID=" & mlng打印ID, "价格等级=" & strPriceGrade, "PrintEmpty=0", str缴款, str找补, strFormat, 2)
                            If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                            If Not mblnPrinted And mobjReport.DataIsEmpty = False Then Exit For '109708
                        End If
                        If mobjReport.DataIsEmpty Then   '109708
                            strClearNOs = strClearNOs & "," & mstrPrintNO
                        Else
                            If i < UBound(Split(strNos, ",")) Then '取下一票据号,供BeforePrint中使用
                                If gblnStrictCtrl Then
                                    mstrInvoice = GetNextBill(mlng领用ID)   '不够时,返回空
                                Else
                                    mstrInvoice = zlStr.Increase(mstrInvoice)
                                End If
                                '票据严格控制时再取一次
                                If mstrInvoice = "" And gblnStrictCtrl Then
                                    If zlCheckInvoiceValied(mlng领用ID, 1, mstrInvoice, mlngShareUseID, mstrUseType) Then
                                        mstrInvoice = GetNextBill(mlng领用ID)   '不够时,返回空
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                '中途失败处理
                If i < UBound(Split(strNos, ",")) + 1 Then '注意：mobjReport_BeforePrint中的提示在这之前
                    If i = 0 Then
                        MsgBox "单据[" & strNos & "]一张也没有打印!" & vbCrLf & _
                            "请在指定新的票据号后使用重打功能打印！", vbInformation, gstrSysName
                    Else
                        MsgBox "单据[" & strNos & "]只打印了前" & i & "张!" & vbCrLf & _
                            "剩下的请在指定新的票据号后使用重打功能打印！", vbInformation, gstrSysName
                    End If
                    For j = i To UBound(Split(strNos, ","))
                        strClearNOs = strClearNOs & "," & Split(strNos, ",")(j)
                    Next
                    strClearNOs = Replace(Mid(strClearNOs, 2), "'", "")
                    GoTo ClearInvoice
                End If
            Else
                '修改多张中的一张时,修改新产生的这张放在最后,因为mobjReport_BeforePrint中要取第一张来收回原来的
                mstrPrintNO = IIf(strAllNOs <> "", strAllNOs & ",", "") & strNos
                If blnNotPrint Then
                    Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp) '调用打印方法但不打印，只生成了票据使用数据
                    If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice '修改时,只清除新单据的开始票据号
                    Call TaxInterface(1, mstrPrintNO, strAllNOs)
                Else
                   '票据接口
                    If BillPrint(1, mstrPrintNO, strAllNOs, "", strClearNOs) = False Then: GoTo ClearInvoice
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                        "发票号=FactNO", "NO=" & mstrPrintNO, "打印ID=" & mlng打印ID, "价格等级=" & strPriceGrade, "PrintEmpty=0", str缴款, str找补, strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                    If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice
                End If
            End If
        Case 2, 4 '重打
            mstrPrintNO = strNos
            If gTy_Module_Para.byt票据分配规则 <> 0 And (mstrReclaimInvoice <> "" Or mbytInFun = 4) Then
                '1.根据系统预定规则打印票据;2-根据用户自定义分配票号
                '回收发票不能为空时,才能按新方式重打票据
                '先分配票号:
                If gblnStrictCtrl Then
                    str发票号 = mstrReclaimInvoice
                    '模拟计算，把需要的票据张数算出来
                    '1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
                    If zlExeCuteBillNoSplit(True, IIf(mbytInFun = 4, 2, 3), mlng领用ID, mstrPrintNO, 0, mstrInvoice, mdatFeeDate, _
                                        1, str发票号, int票据张数) = False Then Exit Sub
                    If zlCheckInvoiceValied(mlng领用ID, int票据张数, mstrInvoice, mlngShareUseID, mstrUseType, _
                                    mlngRestNum, mlngNextUseID, mstrNextInvoice) = False Then Exit Sub
                End If
                str发票号 = mstrReclaimInvoice
                '1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
                If zlExeCuteBillNoSplit(False, IIf(mbytInFun = 4, 2, 3), mlng领用ID, mstrPrintNO, 0, mstrInvoice, mdatFeeDate, _
                                     1, str发票号, int票据张数, mlngNextUseID, mstrNextInvoice) = False Then Exit Sub
                mstrReclaimInvoice = str发票号
                mbln分配票号 = True
                If int票据张数 = 0 Then Exit Sub
               
                '不严格控制票据时保存当前收费票据号
                If Not gblnStrictCtrl And mstrReclaimInvoice <> "" Then
                    var发票号 = Split(mstrReclaimInvoice, ",")
                    zlDatabase.SetPara "当前收费票据号", var发票号(UBound(var发票号)), glngSys, 1121
                End If
                
                If blnNotPrint Then
                     Call TaxInterface(2, mstrPrintNO, strAllNOs)      '打印税控票据
                     ''调用医保重打接口
                     If InsureReprint(blnVirtualPrint, Replace(Split(strNos, ",")(0), "'", ""), lngBalanceID, blnDelRecord, strInvoice) = False Then Exit Sub
                Else
                    '票据接口
                     If BillPrint(2, mstrPrintNO, strAllNOs, strInvoice, strClearNOs) = False Then Exit Sub
                     '调用报表
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                    "发票号=" & str发票号, "NO='NO'", "打印ID=" & mlng打印ID, "价格等级=" & strPriceGrade, "PrintEmpty=0", "", "", strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                End If
                Exit Sub
            End If
            
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
                Call TaxInterface(2, mstrPrintNO, strAllNOs)      '打印税控票据
                    ''调用医保重打接口
                    If InsureReprint(blnVirtualPrint, Replace(Split(strNos, ",")(0), "'", ""), lngBalanceID, blnDelRecord, strInvoice) = False Then Exit Sub
            Else
                If bln分别打印 And UBound(Split(strNos, ",")) > 0 And strAllNOs = "" Then
                    For i = 0 To UBound(Split(strNos, ","))
                        mblnPrinted = False '须在这里初始，因为可能打印部件在调BeforePrint之前就出错返回了
                        mstrPrintNO = Split(strNos, ",")(i)
                        blnPrint = True
                        If gTy_Module_Para.bln工本费 Then
                            '一张单据只有工本费不打印
                            If BillOnlyFactMoney(Replace(mstrPrintNO, "'", "")) Then blnPrint = False
                        End If
                        If blnPrint Then
                            If blnNotPrint Then
                                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                                If Not mblnPrinted Then Exit For
                                Call TaxInterface(3, mstrPrintNO, "")
                            Else
                                If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit For
                                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                                    "发票号=FactNO", "NO=" & mstrPrintNO, "打印ID=" & mlng打印ID, "价格等级=" & strPriceGrade, "PrintEmpty=0", "", "", strFormat, 2)
                                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                                If Not mblnPrinted And mobjReport.DataIsEmpty = False Then Exit For '109708
                            End If
                            If mobjReport.DataIsEmpty = False Then '109708
                                If i < UBound(Split(strNos, ",")) Then '取下一票据号,供BeforePrint中使用
                                    If gblnStrictCtrl Then
                                        mstrInvoice = GetNextBill(mlng领用ID)   '不够时,返回空
                                    Else
                                        mstrInvoice = zlStr.Increase(mstrInvoice)
                                    End If
                                    '票据严格控制时再取一次
                                    If mstrInvoice = "" And gblnStrictCtrl Then
                                        If zlCheckInvoiceValied(mlng领用ID, 1, mstrInvoice, mlngShareUseID, mstrUseType) Then
                                            mstrInvoice = GetNextBill(mlng领用ID)   '不够时,返回空
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next i
                Else
                    '票据接口
                    If BillPrint(2, mstrPrintNO, "", strInvoice, strClearNOs) = False Then Exit Sub
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                        "发票号=FactNO", "NO=" & mstrPrintNO, "打印ID=" & mlng打印ID, "价格等级=" & strPriceGrade, "PrintEmpty=0", "", "", strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                    If Not mblnPrinted Then Exit Sub
                End If
            End If
        Case 3  '退费
            If gTy_Module_Para.byt票据分配规则 <> 0 And mstrReclaimInvoice <> "" Then
                '1.根据系统预定规则打印票据;2-根据用户自定义分配票号
                '回收发票不能为空时,才能按新方式重打票据
                '先分配票号:
                If gblnStrictCtrl Then
                    str发票号 = mstrReclaimInvoice
                    '模拟计算，把需要的票据张数算出来
                    '1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
                    If zlExeCuteBillNoSplit(True, 4, mlng领用ID, strNos, 0, mstrInvoice, mdatFeeDate, 1, str发票号, int票据张数) = False Then Exit Sub
                    If zlCheckInvoiceValied(mlng领用ID, int票据张数, mstrInvoice, mlngShareUseID, mstrUseType) = False Then Exit Sub
                End If
                str发票号 = mstrReclaimInvoice
                '1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
                If zlExeCuteBillNoSplit(False, 4, mlng领用ID, strNos, 0, mstrInvoice, mdatFeeDate, 1, _
                                    str发票号, int票据张数, mlngNextUseID, mstrNextInvoice) = False Then Exit Sub
                mstrReclaimInvoice = str发票号
                mbln分配票号 = True
                If int票据张数 = 0 Then Exit Sub
                
                '不严格控制票据时保存当前收费票据号
                If Not gblnStrictCtrl And mstrReclaimInvoice <> "" Then
                    var发票号 = Split(mstrReclaimInvoice, ",")
                    zlDatabase.SetPara "当前收费票据号", var发票号(UBound(var发票号)), glngSys, 1121
                End If
                
                If blnNotPrint Then
                     Call TaxInterface(3, strNos, "")       '打印税控票据
                Else
                    '票据接口
                  If BillPrint(3, strNos, "", strInvoice, "") = False Then Exit Sub
                     '调用报表
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                    "发票号=" & str发票号, "NO='NO'", "打印ID=" & mlng打印ID, "价格等级=" & strPriceGrade, "PrintEmpty=0", "", "", strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                End If
                Exit Sub
            End If
            If bln分别打印 And UBound(Split(strNos, ",")) > 0 And strAllNOs = "" Then
                For i = 0 To UBound(Split(strNos, ","))
                    mblnPrinted = False '须在这里初始，因为可能打印部件在调BeforePrint之前就出错返回了
                    mstrPrintNO = Split(strNos, ",")(i)
                    blnPrint = True
                    If gTy_Module_Para.bln工本费 Then
                        '一张单据只有工本费不打印
                        If BillOnlyFactMoney(Replace(mstrPrintNO, "'", "")) Then blnPrint = False
                    End If
                    If blnPrint Then
                        If blnNotPrint Then
                            Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                            If Not mblnPrinted Then Exit For
                            Call TaxInterface(3, mstrPrintNO, "")
                        Else
                            If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit For
                            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                                "发票号=FactNO", "NO=" & mstrPrintNO, "打印ID=" & mlng打印ID, "价格等级=" & strPriceGrade, "PrintEmpty=0", "", "", strFormat, 2)
                            If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                            If Not mblnPrinted And mobjReport.DataIsEmpty = False Then Exit For '109708
                        End If
                        If mobjReport.DataIsEmpty = False Then '109708
                            If i < UBound(Split(strNos, ",")) Then '取下一票据号,供BeforePrint中使用
                                If gblnStrictCtrl Then
                                    mstrInvoice = GetNextBill(mlng领用ID)   '不够时,返回空
                                Else
                                    mstrInvoice = zlStr.Increase(mstrInvoice)
                                End If
                                '票据严格控制时再取一次
                                If mstrInvoice = "" And gblnStrictCtrl Then
                                    If zlCheckInvoiceValied(mlng领用ID, 1, mstrInvoice, mlngShareUseID, mstrUseType) Then
                                        mstrInvoice = GetNextBill(mlng领用ID)   '不够时,返回空
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            Else
                mstrPrintNO = strNos
                If blnNotPrint Then
                    Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                    If Not mblnPrinted Then Exit Sub
                    Call TaxInterface(3, mstrPrintNO, "")
                Else
                    If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit Sub
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, _
                        "发票号=FactNO", "NO=" & mstrPrintNO, "打印ID=" & mlng打印ID, "价格等级=" & strPriceGrade, "PrintEmpty=0", "", "", strFormat, 2)
                        If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                    If Not mblnPrinted Then Exit Sub
                End If
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
'    Exit Sub
ClearInvoice:
    On Error GoTo errH
    
    If strClearNOs = "" Then Exit Sub
    strClearNOs = Replace(strClearNOs, "'", "")
    If Left(strClearNOs, 1) = "," Then strClearNOs = Mid(strClearNOs, 2)
    
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

