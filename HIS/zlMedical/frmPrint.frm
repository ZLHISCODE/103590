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

Private mbytInFun As Byte               '1-新单打印,2-重打
Private mlng领用ID As Long              '上次领用ID
Private mstrPrintNO As String           '结帐单据号
Private mlngBalanceID As Long           '结帐ID
Private mstrInvoice As String           '开始票据号
Private mdateBalance As Date            '结帐或重打的时间
Private mblnPrinted As Boolean          '打印票据数据生成是否成功
Private mbytKind As Byte

Private Sub Form_Load()
    Set mobjReport = New clsReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
    mbytInFun = 0
    mlng领用ID = 0
    mstrPrintNO = ""
    mlngBalanceID = 0
    mstrInvoice = ""
    mdateBalance = CDate(0)
    mblnPrinted = False
End Sub


Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrInvoice As Variant)
    Dim strSQL As String, i As Integer, strInvoices As String
    
    '没有票据号,严格控制票据时不打印,不严格控制票据时只打印不处理票据数据
    If mstrInvoice = "" Then
        Cancel = gblnStrictCtrl
        mblnPrinted = Not gblnStrictCtrl
        Exit Sub
    End If
    
    mblnPrinted = False
    '1.严格控制票据时，根据实际的票据张数,重新检查领用ID和票据号
    If gblnStrictCtrl Then
        mlng领用ID = GetInvoiceGroupID(mbytKind, TotalPages, mlng领用ID, glngShareUseID, mstrInvoice)
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case -1
                    MsgBox "单据[" & mstrPrintNO & "]共需要" & TotalPages & "张票据!" & vbCrLf & _
                        "你没有足够的自用和共用的票据,请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "单据[" & mstrPrintNO & "]共需要" & TotalPages & "张票据!" & vbCrLf & _
                        "你没有足够的的共用票据,请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -3
                    MsgBox "单据[" & mstrPrintNO & "]共需要" & TotalPages & "张票据!" & vbCrLf & _
                        "票据号[" & mstrInvoice & "]不在可用领用批次的有效票据号范围内！" & _
                        "请重新输入有效的票据号后重打该单据！", vbInformation, gstrSysName
                Case -4
                    MsgBox "单据[" & mstrPrintNO & "]共需要" & TotalPages & "张票据!" & vbCrLf & _
                        "票据号[" & mstrInvoice & "]所在的领用批次没有足够的票据！" & _
                        "请先打印其它票据,用完当前领用批次后,重打该单据！", vbInformation, gstrSysName
                Case Else
                    MsgBox "票据领用信息访问失败！将来，你可以重打单据[" & mstrPrintNO & "]", vbInformation, gstrSysName
            End Select
            Cancel = True: Exit Sub
        End If
    End If
    
    '2.产生票据使用数据
    On Error GoTo errH
    Select Case mbytInFun
        Case 1
            strSQL = "zl_病人结帐票据_Insert('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng领用ID) & _
                ",'" & UserInfo.姓名 & "',To_Date('" & Format(mdateBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & TotalPages & ")"
        
        Case 2
            strSQL = "zl_病人结帐记录_RePrint('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng领用ID) & _
                ",'" & UserInfo.姓名 & "'," & TotalPages & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSQL, "票据数据生成")
    mblnPrinted = True
    
    '3.传递所用的票据号信息
    For i = 1 To TotalPages
        strInvoices = strInvoices & "," & mstrInvoice
        mstrInvoice = IncStr(mstrInvoice)
    Next
    strInvoices = Mid(strInvoices, 2)
    arrInvoice = Split(strInvoices, ",")
        
    '不严格控制票据时保存到注册表
    If Not gblnStrictCtrl Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "当前结帐票据号", mstrInvoice
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub

Public Sub ReportPrint(ByVal bytInfun As Byte, ByVal strNO As String, ByVal lngBalanceID As Long, _
                        ByRef lngLastUseID As Long, ByVal strInvoice As String, Optional ByVal dateBalance As Date, _
                        Optional str缴款 As String, Optional str找补 As String, Optional ByVal bytKind As Byte = 3)
'参数： bytInfun        =   1-新单打印,2-重打
'       strNO           =   结帐单据号,不带引号
'       lngBalanceID    =   结帐ID
'       lngLastUseID   <=>  最近使用的领用批次ID,初次时为0
'       strInvoice      =   开始票据号，不带引号,不严格控制票据时允许传入空,严格控制时调用前已前检查不能为空
'       dateBalance     =   结帐时间,仅新单打印才传入
    Dim strReportNO As String, strSQL As String
    
    '1.变量传递
    mbytInFun = bytInfun
    mstrPrintNO = strNO
    mlngBalanceID = lngBalanceID
    mlng领用ID = lngLastUseID
    mstrInvoice = strInvoice
    mdateBalance = dateBalance
    strReportNO = "ZL1_BILL_1862"
    mbytKind = bytKind
    
    '2.打印调用
    Select Case mbytInFun
        Case 1  '新单打印

            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "结帐ID=" & mlngBalanceID, str缴款, str找补, 2)
            If Not mblnPrinted Then GoTo ClearInvoice

        Case 2  '重打

            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "结帐ID=" & mlngBalanceID, "", "", 2)
            If Not mblnPrinted Then Exit Sub

    End Select
    
    '3.传回最近使用的领用ID
    lngLastUseID = mlng领用ID
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
