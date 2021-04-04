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
Private mobjFactProperty As clsFactProperty
Private mobjInvoice As clsInvoice
Private mbytInFun As Byte               '1-新单打印,2-重打,3-红票打印
Private mlng领用ID As Long              '上次领用ID
Private mstrPrintNO As String           '结帐单据号
Private mlngBalanceID As Long           '结帐ID
Private mstrInvoice As String           '开始票据号
Private mdateBalance As Date            '结帐或重打的时间
Private mblnPrinted As Boolean          '打印票据数据生成是否成功
Private mblnInitInvoice As Boolean

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

 

Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
    Set mobjFactProperty = Nothing
    Set mobjInvoice = Nothing
    
    mbytInFun = 0
    mlng领用ID = 0
    mstrPrintNO = ""
    mlngBalanceID = 0
    mstrInvoice = ""
    mdateBalance = CDate(0)
    mblnPrinted = False
    mblnInitInvoice = False
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
    Dim cllPro As Collection
    Dim strUserType As String, bytKind As Byte '0:住院医疗费收据,1-门诊医疗费收据
    
    With mPrintSheet
        .blnCalcMoney = True
        .lngPrePage = 0
        .lngCurPrintRow = 0
        .blnUseOnlyOneInvoice = False
    End With
    
    '没有票据号,严格控制票据时不打印,不严格控制票据时只打印不处理票据数据
    If mblnInitInvoice = False Then
        mobjInvoice.zlInitCommon glngSys, gcnOracle, gstrDBUser
        mblnInitInvoice = True
    End If
    If mstrInvoice = "" Then
        Cancel = mobjFactProperty.严格控制
        mblnPrinted = Not mobjFactProperty.严格控制
        mPrintSheet.blnCalcMoney = False '不计算票据金额
        Exit Sub
    End If
    Set cllPro = New Collection
    strUserType = ""
    If mobjFactProperty.使用类别 <> "" Then strUserType = "(" & mobjFactProperty.使用类别 & ")"
    mblnPrinted = False
    '1.严格控制票据时，根据实际的票据张数,重新检查领用ID和票据号
    If mobjFactProperty.严格控制 Then
        If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.姓名, mobjFactProperty.票种, _
            mobjFactProperty.使用类别, mlng领用ID, mobjFactProperty.共享批次ID, mlng领用ID, TotalPages, mstrInvoice) = False Then
            Cancel = True: Exit Sub
        End If
       If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case -1
                    MsgBox IIf(mbytInFun = 3, "本次退费发票(红票)打印", "单据[" & mstrPrintNO & "]") & "需要" & TotalPages & "张票据!" & vbCrLf & _
                        "你没有足够的自用和共用的票据" & strUserType & ",请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -2
                    MsgBox IIf(mbytInFun = 3, "本次退费发票(红票)打印", "单据[" & mstrPrintNO & "]") & "需要" & TotalPages & "张票据!" & vbCrLf & _
                        "你没有足够的的共用票据,请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -3
                    MsgBox IIf(mbytInFun = 3, "本次退费发票(红票)打印", "单据[" & mstrPrintNO & "]") & "需要" & TotalPages & "张票据!" & vbCrLf & _
                        "票据号[" & mstrInvoice & "]不在可用领用批次的有效票据号范围内！" & _
                        "请重新输入有效的票据号后重打该单据！", vbInformation, gstrSysName
                Case -4
                    MsgBox IIf(mbytInFun = 3, "本次退费发票(红票)打印", "单据[" & mstrPrintNO & "]") & "需要" & TotalPages & "张票据!" & vbCrLf & _
                        "票据号[" & mstrInvoice & "]所在的领用批次没有足够的票据！" & _
                        "请先打印其它票据,用完当前领用批次后,重打该单据！", vbInformation, gstrSysName
                Case Else
                    MsgBox "票据领用信息访问失败！将来，你可以" & IIf(mbytInFun = 3, "重打该单据！", "重打单据[" & mstrPrintNO & "]！"), vbInformation, gstrSysName
            End Select
            Cancel = True: Exit Sub
        End If
    End If
    
    '2.产生票据使用数据
    bytKind = IIf(mobjFactProperty.票种 = 3, 0, 1)
    On Error GoTo errH
    Select Case mbytInFun
        Case 1
            strSQL = "zl_病人结帐票据_Insert('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng领用ID) & _
                ",'" & UserInfo.姓名 & "',To_Date('" & Format(mdateBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & TotalPages & "," & bytKind & ")"
        
        Case 2
            strSQL = "zl_病人结帐记录_RePrint('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng领用ID) & _
                ",'" & UserInfo.姓名 & "'," & TotalPages & "," & bytKind & ")"
        Case 3 '红票打印
            'Zl_病人结帐记录_Reprint
            strSQL = "Zl_病人结帐记录_Reprint("
            '  No_In       病人预交记录.No%Type,
            strSQL = strSQL & "'" & mstrPrintNO & "',"
            '  票据号_In   票据使用明细.号码%Type,
            strSQL = strSQL & "'" & mstrInvoice & "',"
            '  领用id_In   票据使用明细.领用id%Type,
            strSQL = strSQL & "" & ZVal(mlng领用ID) & ","
            '  使用人_In   票据使用明细.使用人%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  票据张数_In Number,
            strSQL = strSQL & "" & TotalPages & ","
            '  票种_In     Number := 0, --0:住院医疗费收据,1-门诊医疗费收据
            strSQL = strSQL & bytKind & ","
            '  红票打印_In Number := 0, --0:正常重打,1-作废时候红票打印
            strSQL = strSQL & "" & 1 & ","
            '  使用时间_In Date:=Null
            strSQL = strSQL & "" & "To_date('" & Format(mdateBalance, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')" & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSQL, "票据数据生成")
    mblnPrinted = True
    
    '3.传递所用的票据号信息
    For i = 1 To TotalPages
        strInvoices = strInvoices & "," & mstrInvoice
        If i < TotalPages Then mstrInvoice = zlCommFun.IncStr(mstrInvoice)
    Next
    strInvoices = Mid(strInvoices, 2)
    If strInvoices <> "" Then arrInvoice = Split(strInvoices, ",")
    
    mPrintSheet.arrInvoice = arrInvoice
    
    '不严格控制票据时保存到注册表
    If Not mobjFactProperty.严格控制 Then
        zlDatabase.SetPara "当前结帐票据号", mstrInvoice, glngSys, 1137
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub
Public Sub ReportPrint(ByVal bytInfun As Byte, ByVal strNO As String, ByVal lngBalanceID As Long, _
                        ByRef objFactProperty As clsFactProperty, _
                        ByVal strInvoice As String, Optional ByVal dateBalance As Date, _
                        Optional str缴款 As String, Optional str找补 As String, Optional lngPatientID As Long, _
                        Optional intLocalFormat As Integer, Optional blnPrintBillEmpty As Boolean = False, _
                        Optional blnInsurePrint As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:结帐票据打印
    '入参:bytInfun:1-新单打印,2-重打,3-红票打印
    '       strNO:结帐单据号,不带引号
    '       lngBalanceID:结帐ID
    '       objFactProperty-发票属性控制
    '       lngLastUseID:最近使用的领用批次ID,初次时为0
    '       lngShareUseID:共享批次
    '       strUseType:使用类别
    '       strInvoice:开始票据号，不带引号,不严格控制票据时允许传入空,严格控制时调用前已前检查不能为空
    '       dateBalance :结帐时间,仅新单打印才传入
    '       lngPatientID:合约单位结帐按病人分别打印,每次打印传入当前病人ID
    '       intLocalFormat:按指定的格式打印
    '       blnInsurePrint:是否医保接口打印
    '出参:
    '       blnPrintBillEmpty-是否打印空票据(55052)
    '返回:
    '编制:刘兴洪
    '日期:2011-05-03 17:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strReportNO As String, strSQL As String, strFormat As String
    Dim arrInvoice As Variant
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    If mobjInvoice Is Nothing Then Set mobjInvoice = New clsInvoice:  mblnInitInvoice = False
    
    If mblnInitInvoice = False Then
        mobjInvoice.zlInitCommon glngSys, gcnOracle, gstrDBUser
        mblnInitInvoice = True
    End If
    
    blnPrintBillEmpty = False
    '1.变量传递
    mbytInFun = bytInfun: mstrPrintNO = strNO
    mlngBalanceID = lngBalanceID: mlng领用ID = objFactProperty.LastUseID
    mstrInvoice = strInvoice: mdateBalance = dateBalance
    Set mobjFactProperty = objFactProperty
    
    If objFactProperty.票种 = 3 Then
        If mbytInFun = 3 Then
            strReportNO = "ZL" & glngSys \ 100 & "_BILL_1137_5"
        Else
            strReportNO = "ZL" & glngSys \ 100 & "_BILL_1137"
        End If
    Else
        If mbytInFun = 3 Then
            strReportNO = "ZL" & glngSys \ 100 & "_BILL_1137_6"
        Else
            strReportNO = "ZL" & glngSys \ 100 & "_BILL_1137_2"
        End If
    End If
    '选择的打印格式
    strFormat = IIf(intLocalFormat <= 0, "", "ReportFormat=" & intLocalFormat)
    mblnPrinted = False
    
    '2.打印调用
    Select Case mbytInFun
        Case 1  '新单打印
            If Not gobjTax Is Nothing And gblnTax Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)   '调用打印方法但不打印，只生成了票据使用数据
                If IsArray(arrInvoice) Then
                    mstrInvoice = arrInvoice(0)
                Else
                    mstrInvoice = arrInvoice
                End If
                If Not mblnPrinted Then GoTo ClearInvoice
                
                If Not gobjTax Is Nothing And gblnTax Then
                    gstrTax = gobjTax.zlTaxInPrint(gcnOracle, mlngBalanceID)
                    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
                End If
            Else
                If gblnBillPrint Then
                    If gobjBillPrint.zlPrintBill("", mlngBalanceID) = False Then GoTo ClearInvoice
                End If
                If blnInsurePrint Then
                    Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)   '调用打印方法但不打印，只生成了票据使用数据
                Else
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "结帐ID=" & mlngBalanceID, "病人ID=" & lngPatientID, "PrintEmpty=0", str缴款, str找补, strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                    If Not mblnPrinted Then GoTo ClearInvoice
                End If
            End If
        Case 2  '重打
            If Not gobjTax Is Nothing And gblnTax Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)
                If IsArray(arrInvoice) Then
                    mstrInvoice = arrInvoice(0)
                Else
                    mstrInvoice = arrInvoice
                End If
                If Not mblnPrinted Then Exit Sub
                
                If Not gobjTax Is Nothing And gblnTax Then
                    MsgBox "请在准备好之后按确定开始打印。", vbInformation, gstrSysName
                    gstrTax = gobjTax.zlTaxInReput(gcnOracle, mlngBalanceID)
                    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
                End If
            Else
                If gblnBillPrint Then
                    If gobjBillPrint.zlRePrintBill("", mlngBalanceID, strInvoice) = False Then Exit Sub
                End If
                
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "结帐ID=" & mlngBalanceID, "病人ID=" & lngPatientID, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
        Case 3 '红票打印
            If Not gobjTax Is Nothing And gblnTax Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)
                If IsArray(arrInvoice) Then
                    mstrInvoice = arrInvoice(0)
                Else
                    mstrInvoice = arrInvoice
                End If
                If Not mblnPrinted Then Exit Sub
                
                If Not gobjTax Is Nothing And gblnTax Then
                    MsgBox "请在准备好之后按确定开始打印。", vbInformation, gstrSysName
                    gstrTax = gobjTax.zlTaxInReput(gcnOracle, mlngBalanceID)
                    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
                End If
            Else
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "结帐ID=" & mlngBalanceID, "病人ID=" & lngPatientID, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
    End Select
    '3.传回最近使用的领用ID
    mobjFactProperty.LastUseID = mlng领用ID
    Exit Sub
    
ClearInvoice:
    On Error GoTo errH
    strSQL = "Zl_票据起始号_Update('" & strNO & "','',3)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errH:
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
    strSQL = strSQL & "" & mobjFactProperty.票种 & ")"
    zlDatabase.ExecuteProcedure strSQL, "更新票据金额"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
