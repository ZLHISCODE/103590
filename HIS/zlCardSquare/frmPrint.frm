VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "票据打印"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mstrPrintNO As String           '要打印的单据号
Private mstrInvoice As String           '要打印的票据号
Private mEditType As gCardType          '操作类型
Private mlng领用ID As Long              '发卡领用ID
Private mstrUseType As String           '使用类别
Private mdtPrintdate As Date            '打印时间
Private mUserName As String             '使用人
                                
Public Sub PrintBill(ByVal strNO As String, ByVal strCardNo As String, _
                     ByVal strInvoice As String, ByVal lngCardTypeID As Long, ByVal blnPrint As Boolean, _
                     ByVal EditType As gCardType, ByVal bytPrintFormat As Byte, ByVal lng领用ID As Long, _
                     ByVal strUseType As String, ByVal dtPrintdate As Date, ByVal UserName As String, _
                     Optional blnPrepayPrint As Boolean = False, Optional strPrePayNo As String = "", _
                     Optional lng预交病人ID As Long = 0, Optional dat预交时间 As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行发卡票据打印
    '参数：strNO           单据号
    '      strPrePayNo     预交单号
    '      strCardNo       卡号
    '      lng预交病人ID   预交病人ID
    '      strInvoice      票据号
    '      lngCardTypeID   卡类别ID
    '      blnPrint        是否打印
    '      blnPrepayPrint  是否打印预交单
    '      EditType        操作类型
    '      bytPrintFormat  打印格式:发卡|绑定卡
    '      lng领用ID       发卡领用ID
    '      strUseType      使用类别
    '      dtPrintdate     打印时间
    '      UserName        使用人
    '      blnPrepayPrint  是否打印预交单
    '      strPrePayNo     预交单号
    '编制:李南春
    '日期:2014-04-10 13:41:24
    '问题号:57950
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFormat As String
    On Error GoTo Errhand
    mstrPrintNO = strNO
    mstrInvoice = strInvoice
    mlng领用ID = lng领用ID
    mstrUseType = strUseType
    mdtPrintdate = dtPrintdate
    mUserName = UserName
    
    If blnPrepayPrint Then
        '打印预交票据
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & strPrePayNo, "病人ID=" & lng预交病人ID, "收款时间=" & Format(dat预交时间, "yyyy-mm-dd HH:MM:SS"), 2)
    End If
    
    If Not blnPrint Then Exit Sub
    strFormat = IIf(bytPrintFormat = 0, "", "ReportFormat=" & bytPrintFormat)
    
    If mEditType = Cr_绑定卡 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1107", Me, "卡类别ID=" & lngCardTypeID, "NO=" & strCardNo, "卡号=" & strCardNo, "PrintEmpty=0", strFormat, 2)
    ElseIf gbln收费发票 Then
        Set mobjReport = New clsReport
        Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1107", Me, "卡类别ID=" & lngCardTypeID, "NO=" & strNO, "卡号=" & strCardNo, "PrintEmpty=0", strFormat, 2)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1107", Me, "卡类别ID=" & lngCardTypeID, "NO=" & strNO, "卡号=" & strCardNo, "PrintEmpty=0", strFormat, 2)
    End If
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
         Resume
    End If
    SaveErrLog
End Sub

Private Sub Form_Load()
    On Error GoTo Errhand
    
    Set mobjReport = New clsReport
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
         Resume
    End If
    SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Errhand
    
    Set mobjReport = Nothing
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
         Resume
    End If
    SaveErrLog
End Sub

Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant)
    Dim lng领用ID As Long
    Dim strSQL As String
    arrBill = Split(mstrInvoice, ",")
    On Error GoTo errH
    If gblnBill发卡 Then
        lng领用ID = GetInvoiceGroupID(1, TotalPages, mlng领用ID, glngShareUseID, mstrInvoice, mstrUseType)
        If lng领用ID <= 0 Then
            Select Case lng领用ID
                Case -1
                    MsgBox "单据[" & mstrPrintNO & "]" & "共需要" & TotalPages & "张票据！" & vbCrLf & _
                        "你没有足够的自用和共用的票据，请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "单据[" & mstrPrintNO & "]" & "共需要" & TotalPages & "张票据！" & vbCrLf & _
                        "你没有足够的的共用票据，请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -3
                    MsgBox "单据[" & mstrPrintNO & "]" & "共需要" & TotalPages & "张票据！" & vbCrLf & _
                        "票据号[" & mstrInvoice & "]不在可用领用批次的有效票据号范围内！" & _
                        "请重新输入有效的票据号后重打该单据！", vbInformation, gstrSysName
                Case -4
                    MsgBox "单据[" & mstrPrintNO & "]" & "共需要" & TotalPages & "张票据！" & vbCrLf & _
                        "票据号[" & mstrInvoice & "]所在的领用批次没有足够的票据！" & _
                        "请先打印其它票据,用完当前领用批次后，重打该单据！", vbInformation, gstrSysName
                Case Else
                    MsgBox "票据领用信息访问失败！将来，你可以重打单据[" & mstrPrintNO & "]！", vbInformation, gstrSysName
            End Select
            Cancel = True: Exit Sub
        End If
    End If
    strSQL = "Zl_病人发卡票据_Print("
    '  No_In           Varchar2,
    strSQL = strSQL & "'" & Replace(mstrPrintNO, "'", "") & "'" & ","
    '  票据号_In       票据使用明细.号码%Type,
    strSQL = strSQL & "'" & mstrInvoice & "',"
    '  领用id_In       票据使用明细.领用id%Type,
    strSQL = strSQL & "" & ZVal(lng领用ID) & ","
    '  使用人_In       票据使用明细.使用人%Type,
    strSQL = strSQL & "'" & mUserName & "',"
    '  使用时间_In     票据使用明细.使用时间%Type,
    strSQL = strSQL & "To_Date('" & Format(mdtPrintdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  操作类型_In     Number
    strSQL = strSQL & IIf(mEditType = Cr_换卡, 5, 1) & ","
    '  票据张数_In     Number := 1,
    strSQL = strSQL & "" & TotalPages & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "票据数据生成")
    
    '不严格控制票据时保存到注册表
    '更新本地票据
    If Not gblnBill发卡 Then
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

Public Sub PrintReBill(ByVal strSelect As String, ByVal strCardNo As String, ByVal lngCardTypeID As Long, ByVal bytPrintPayCard As Byte)
    '功能:重打票据(旧模式)
    Dim strFormat As String
    On Error GoTo errH
    mstrInvoice = ""
  
    If strSelect = "发卡" Then
        If strCardNo = "" Then ShowMsgbox "没选中相关的医疗卡": Exit Sub
        strFormat = IIf(bytPrintPayCard = 0, "", "ReportFormat=" & bytPrintPayCard)
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1107", Me, "卡类别ID=" & lngCardTypeID, "NO=" & strCardNo, "卡号=" & strCardNo, "缴款=" & 0, "找补=" & 0, "PrintEmpty=0", strFormat, 2)
    Else
        '预存款打印
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function RePrintBill(ByVal frmParent As Object, ByVal strCardNo As String, ByVal lngCardTypeID As Long, _
                             ByVal strUseType As String, ByVal strPrintNo As String, ByVal intPrintMode As Integer, _
                             ByVal bytPrintPayCard As Byte, Optional ByVal bln重打 As Boolean) As Boolean
    '--------------------------------------------------------------------------------------------------------------------------------------------
    '功能:当前收款记录重新打印一张票据(新模式)
    '入参:   strUseType-使用类别

    '        blnVirtualPrint-医保接口内调用打印，HIS只走票号不实际打印
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-19 17:18:19
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    Dim strInvoice As String
    Dim blnValid As Boolean, blnInput As Boolean
    Dim lng领用ID As Long, strBackInvoice As String
    Dim blnReprint As Boolean, strFormat As String
    
    On Error GoTo errH
    '如果严格控制票据使用
    If gblnBill发卡 Then
        If bln重打 Then
            lng领用ID = CheckUsedBill(1, glngShareUseID, , strUseType)
            Select Case lng领用ID
                Case -1
                    MsgBox "你没有自用和共用的挂号票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            If lng领用ID <= 0 Then Exit Function
        End If
        If intPrintMode = 3 Then
            '获取收回票据
            strSQL = _
            "   Select A.号码" & vbNewLine & _
            "   From 票据使用明细 A" & vbNewLine & _
            "   Where A.性质 = 1 And a.原因 <> 6 " & vbNewLine & _
            "       And A.票种 = 1 And A.打印id = (Select Max(ID) From 票据打印内容 Where 数据性质 = [2] And NO = [1])" & vbNewLine & _
            "   Order By 号码"
            Set rsInvoice = zlDatabase.OpenSQLRecord(strSQL, "获取收回票据", strPrintNo, 5)
            Do While Not rsInvoice.EOF
                strBackInvoice = strBackInvoice & "," & rsInvoice!号码
                rsInvoice.MoveNext
            Loop
            If strBackInvoice <> "" Then strBackInvoice = Mid(strBackInvoice, 2)
        End If
        blnReprint = bln重打
    End If
    
     '取下一个票据号码
    If Not gblnBill发卡 Then
        '有可能是第一次使用
        Do
            blnInput = False
            '非严格控制时直接从本地读取
            strInvoice = zlDatabase.GetPara("当前收费票据号", glngSys, 1121)
            mstrInvoice = strInvoice
            
            If strInvoice = "" Then
                strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定将要使用的开始票据号。" & _
                                vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                "", frmParent.Left + 1500, frmParent.Top + 1500))
                blnInput = True
            Else
                strInvoice = zlCommFun.IncStr(strInvoice)
                strInvoice = UCase(InputBox("请确认重打使用的开始票据号码：", gstrSysName, _
                                strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                blnInput = True
            End If
                
            '用户取消输入,允许打印
            If strInvoice = "" Then
                If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnValid = True
            Else
                '检查输入有效性
                If blnInput Then
                    If zlCommFun.ActualLen(strInvoice) <> gbyt收费 Then
                        MsgBox "输入的票据号码长度应该为 " & gbyt收费 & " 位！", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            End If
        Loop While Not blnValid
        
    Else
        If blnReprint Then
            Do
                '根据票据领用读取
                blnInput = False
                strInvoice = GetNextBill(lng领用ID)
                If strInvoice = "" Then
                    '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
                    strInvoice = UCase(InputBox("无法根据票据领用情况获取将要使用的开始票据号，" & _
                                    vbCrLf & "请你输入将要使用的开始票据号码：", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = UCase(InputBox("请确认重打使用的开始票据号码：", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If
                
                '用户取消输入,不打印
                If strInvoice = "" Then Exit Function
                
                '检查输入有效性
                If blnInput Then
                    If GetInvoiceGroupID(1, 1, lng领用ID, glngShareUseID, strInvoice, strUseType) = -3 Then
                        MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        Else
            strInvoice = ""
        End If
    End If
    
    mlng领用ID = lng领用ID
    mstrInvoice = strInvoice
    '执行数据处理
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    strFormat = IIf(bytPrintPayCard = 0, "", "ReportFormat=" & bytPrintPayCard)
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1107", Me, "卡类别ID=" & lngCardTypeID, "NO=" & strPrintNo, "卡号=" & strCardNo, "PrintEmpty=0", strFormat, 2)
    
    RePrintBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


