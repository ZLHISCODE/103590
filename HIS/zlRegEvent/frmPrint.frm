VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "Ʊ�ݴ�ӡ"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
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
Private mbytInFun As Byte                 '1-�µ���ӡ,2-�ش�,3-�˷Ѵ�ӡ; 4-����Ʊ��;6-�˷�Ʊ��(��Ʊ)��ӡ
Private mlng����ID As Long              '�ϴ�����ID
Private mstrPrintNO As String           'Ҫ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...
Private mstrInvoice As String           '��ʼƱ�ݺ�
Private mdatFeeDate As Date             '���õ������ݵĵǼ�ʱ��
Private mblnPrinted As Boolean          'Ʊ�����������Ƿ�ɹ�(�Ƿ��Ѵ�ӡ)
Private mstrReclaimInvoice As String    'Ҫ����յķ�Ʊ��,��1-����ϵͳԤ���������Ʊ�ź�2-�����û������������Ʊ����Ч
Private mlngShareUseID As Long '��ӡ�Ĺ�������ID
Private mstrUseType As String
Private mbln����Ʊ�� As Boolean

Private Type Ty_PrintSheet
    blnCalcMoney As Boolean '�Ƿ��ۼƷ�Ʊ���
    lngPrePage As Long '��һҳҳ��
    lngGridCount As Long '��ǰҳ�Ѵ�ӡ������
    lngCurPrintRow As Long '��ǰ��ӡ������������ҳ��
    dblInvoiceMoney As Double '��ǰҳ�ۼƷ�Ʊ���
    arrInvoice As Variant '��Ʊ�ţ���ҳ��һһ��Ӧ
    blnUseOnlyOneInvoice As Boolean '�Ƿ��ʹ��һ�ŷ�Ʊ
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
                '�������һҳ������
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
    
    'û��Ʊ�ݺ�,�ϸ����Ʊ��ʱ����ӡ,���ϸ����Ʊ��ʱֻ��ӡ������Ʊ������
    If mstrInvoice = "" Then
        Cancel = gblnBill�Һ�
        mblnPrinted = Not gblnBill�Һ�
        mPrintSheet.blnCalcMoney = False '������Ʊ�ݽ��
        Exit Sub
    End If
    
    If CheckInvoiceValied(TotalPages, mbytInFun = 6) = False Then Cancel = True: Exit Sub
    
    On Error GoTo errH
    '2.����Ʊ������
    Select Case mbytInFun
        Case 1, 4
            strSQL = "Zl_���˹Һ�Ʊ��_Insert("
            '  No_In           Varchar2,
            strSQL = strSQL & "'" & Replace(mstrPrintNO, "'", "") & "'" & ","
            '  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
            strSQL = strSQL & "'" & mstrInvoice & "',"
            '  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
            strSQL = strSQL & "" & ZVal(mlng����ID) & ","
            '  ʹ����_In       Ʊ��ʹ����ϸ.ʹ����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ʹ��ʱ��_In     Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
            strSQL = strSQL & "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  Ʊ������_In     Number := 1,
            strSQL = strSQL & "" & TotalPages & ","
            '  ҽ���ӿڴ�ӡ_In Number := 0,
            strSQL = strSQL & "0,"
            '  �շ�Ʊ��_In Number:=0
            strSQL = strSQL & "" & IIf(gblnSharedInvoice, 1, 0) & ")"
        Case 2, 3
            '����Ƕ��ţ�ֻ��Ҫ��һ�ŵ��ݺž�����(�޸Ķ����е�һ��ʱ,���һ�����µ�)
            strSQL = "Zl_���˹Һż�¼_Reprint('" & Replace(Split(mstrPrintNO, ",")(0), "'", "") & "','" & mstrInvoice & "'," & ZVal(mlng����ID) & ",'" & UserInfo.���� & "'," & _
                    "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & IIf(mbytInFun = 2, "1", "0") & _
                    "," & TotalPages & ",'" & mstrReclaimInvoice & "'," & IIf(gblnSharedInvoice, 1, 0) & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSQL, "Ʊ����������")
    mblnPrinted = True
    
    '3.�������õ�Ʊ�ݺ���Ϣ
    For i = 1 To TotalPages
        strInvoices = strInvoices & "," & mstrInvoice
        If i < TotalPages Then mstrInvoice = zlCommFun.IncStr(mstrInvoice)
    Next
    strInvoices = Mid(strInvoices, 2)
    arrInvoice = Split(strInvoices, ",")
    
    mPrintSheet.arrInvoice = arrInvoice
    
    strSQL = "Zl_ƾ����ӡ��¼_Update(4,'" & mstrPrintNO & "',1,'" & UserInfo.���� & "','��Ʊ��:" & strInvoices & "')"
    zlDatabase.ExecuteProcedure strSQL, "ƾ����ӡ��¼"

    '���ϸ����Ʊ��ʱ���浽ע���
    '���±���Ʊ��
    If Not gblnBill�Һ� Then
        If gblnSharedInvoice Then
            zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", mstrInvoice, glngSys, 1121
        Else
            zlDatabase.SetPara "��ǰ�Һ�Ʊ�ݺ�", mstrInvoice, glngSys, 1111
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

Private Function CheckInvoiceValied(Optional int���� As Integer = 1, _
    Optional ByVal blnDelFeePrint As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鷢Ʊ�Ƿ�Ϸ�(�ϸ����Ʊ��ʱ)
    '���:int���� -��Ҫ�ķ�Ʊ����
    '     blnDelFeePrint-�˷ѷ�Ʊ(��Ʊ)��ӡ
    '����:
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2013-03-27 13:01:41
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gblnBill�Һ� Then CheckInvoiceValied = True: Exit Function
    '1.�ϸ����Ʊ��ʱ������ʵ�ʵ�Ʊ������,���¼������ID��Ʊ�ݺ�
    mlng����ID = GetInvoiceGroupID(IIf(gblnSharedInvoice, 1, 4), int����, mlng����ID, mlngShareUseID, mstrInvoice, mstrUseType)
    '���ݺϷ�
    If mlng����ID > 0 Then CheckInvoiceValied = True: Exit Function
    Select Case mlng����ID
        Case -1
            MsgBox IIf(blnDelFeePrint, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "����Ҫ" & int���� & "��Ʊ�ݣ�" & vbCrLf & _
                "��û���㹻�����ú͹��õ�Ʊ�ݣ�������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
        Case -2
            MsgBox IIf(blnDelFeePrint, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "����Ҫ" & int���� & "��Ʊ�ݣ�" & vbCrLf & _
                "��û���㹻�ĵĹ���Ʊ�ݣ�������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
        Case -3
            MsgBox IIf(blnDelFeePrint, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "����Ҫ" & int���� & "��Ʊ�ݣ�" & vbCrLf & _
                "Ʊ�ݺ�[" & mstrInvoice & "]���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ�" & _
                "������������Ч��Ʊ�ݺź��ش�õ��ݣ�", vbInformation, gstrSysName
        Case -4
            MsgBox IIf(blnDelFeePrint, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "����Ҫ" & int���� & "��Ʊ�ݣ�" & vbCrLf & _
                "Ʊ�ݺ�[" & mstrInvoice & "]���ڵ���������û���㹻��Ʊ�ݣ�" & _
                "���ȴ�ӡ����Ʊ��,���굱ǰ�������κ��ش�õ��ݣ�", vbInformation, gstrSysName
        Case Else
            MsgBox "Ʊ��������Ϣ����ʧ�ܣ������������" & IIf(blnDelFeePrint, "�ش�õ��ݣ�", "�ش򵥾�[" & mstrPrintNO & "]��"), vbInformation, gstrSysName
    End Select
End Function


Private Function InsureReprint(ByVal blnVirtualPrint As Boolean, ByVal strNos As String, _
    ByVal lng����ID As Long, ByVal bln�˷� As Boolean, ByRef strInvoice As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���µ���ҽ����ӡ�ӿ�
    '���:blnVirtualPrint-�Ƿ����ҽ���ӿڴ�ӡ
    '       strNos-���ݺ�
    '       bln�˷�-�Ƿ��˷�
    '       strInvoice-��Ʊ��
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-03-27 17:01:02
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer
    On Error GoTo errHandle
    If Not blnVirtualPrint Then InsureReprint = True: Exit Function
    Call gclsInsure.RePrintBill(intInsure, lng����ID, strInvoice)
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
                        Optional str�ɿ� As String, Optional str�Ҳ� As String, _
                        Optional intPrintFormat As Integer, Optional blnVirtualPrint As Boolean, _
                        Optional ByVal blnDelRecord As Boolean, Optional strUseType As String = "", _
                        Optional blnPrintBillEmpty As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ʊ�ݴ�ӡ,�������
    '���:bytInfun :1-�µ���ӡ,2-�˷Ѵ�ӡ,3-�ش�,4-����Ʊ��(ֻ��:2-��ϵͳԤ�������3-�û��Զ�����ʱ��ת��),6-�˷�Ʊ��(��Ʊ)��ӡ
    '       strNOs - �µ�ʱҪ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...,
    '                   - �޸�ʱ,�����µ��ݺ�,ֻ��һ��,���ڴ�ӡȡ���������ʼƱ�ݺ�
    '                   - �˷�Ʊ��(��Ʊ)��ӡʱ������������
    '       strReclaimInvoice-Ҫ����յķ�Ʊ��,����ö��ŷ���'F0000001','F0000002',...
    '       lngLastUseID-���ʹ�õ���������ID,����ʱΪ0
    '       strInvoice-��ʼƱ�ݺţ���������,���ϸ����Ʊ��ʱ�������,�ϸ����ʱ����ǰ��ǰ��鲻��Ϊ��
    '       datFeeDate-���õ������ݵĵǼ�ʱ��
    '       intPrintFormat-��ӡ��ʽ(��ӡ��ʽ���)
    '       blnVirtualPrint-ҽ���ӿ��ڵ��ô�ӡ��HISֻ��Ʊ�Ų�ʵ�ʴ�ӡ
    '       blnDelRecord-�ش�ʱ���Ƿ��Ƕ��˷Ѽ�¼�����ش�(Ŀǰֻ�б���ҽ��(ҽ���ӿڴ�ӡƱ��)������)
    '       lngShareUseID-��������
    '       strUseType-ʹ�����
    '       lng��ӡID-����Ĵ�ӡID(blnOnePatiPrint=trueʱ����),������Ը��ݴ�ӡID�ӡ���ʱƱ�ݴ�ӡ���ݡ�����ʱ��������ȡ��Ӧ���շѵ���
    '                 ֮����Ҫ��ʱ����Ҫԭ������Ϊ�����˴�ӡʱ�����ݺſ��ܻ���ɣ����������屨����������
    ' ����:
    '   blnPrintBillEmpty-�Ƿ��ӡ�Ŀձ�����()
    '����:���˺�
    '����:2011-04-29 12:01:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
  
    Dim i As Integer, j As Integer, strPrintNO As String, blnPrint As Boolean, blnTrans As Boolean
    Dim strReportNO As String, strSQL As String, strClearNOs As String, strFormat As String, lngBalanceID As Long
    Dim blnNotPrint As Boolean, varTmp As Variant '�ձ�������Ҫ��Ϊ�˴�����ú����ķ���ֵ
    Dim str��Ʊ�� As String, intƱ������ As Integer
    
    Me.Caption = "��ӡ" '����Form_Load�¼�
    blnPrintBillEmpty = False
    mbln����Ʊ�� = False
    '1.��������
    mlngShareUseID = lngShareUseID
    mbytInFun = bytInFun
    mlng����ID = lngLastUseID: mstrUseType = strUseType
    mstrInvoice = strInvoice
    mdatFeeDate = datFeeDate
    mstrReclaimInvoice = strReclaimInvoice
    strReportNO = "ZL" & glngSys \ 100 & "_BILL_1111"
    strFormat = IIf(intPrintFormat = 0, "", "ReportFormat=" & intPrintFormat)
    mstrPrintNO = ""
    mblnPrinted = False
    blnNotPrint = blnVirtualPrint
    '2.��ӡ����
    Select Case mbytInFun
        Case 1 '�µ���ӡ
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp) '���ô�ӡ����������ӡ��ֻ������Ʊ��ʹ������
                If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice '�޸�ʱ,ֻ����µ��ݵĿ�ʼƱ�ݺ�
            Else
                If gblnBillPrint Then
                    On Error Resume Next
                    If gobjBillPrint.zlPrintBill_Reg("'" & strNos & "'") = False Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice
                    On Error GoTo errH
                End If
               'Ʊ�ݽӿ�
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", str�ɿ�, str�Ҳ�, strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice
            End If
        Case 3, 4 '�ش򡢲���
            mstrPrintNO = strNos
            
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
                ''����ҽ���ش�ӿ�
                If InsureReprint(blnVirtualPrint, Replace(Split(strNos, ",")(0), "'", ""), lngBalanceID, blnDelRecord, strInvoice) = False Then Exit Sub
            Else
                If gblnBillPrint Then
                    On Error Resume Next
                    If mbytInFun = 3 Then '�ش����Ʊ������
                        If gobjBillPrint.zlEraseBill_Reg("'" & strNos & "'") = False Then Exit Sub
                    End If
                    If gobjBillPrint.zlRePrintBill_Reg("'" & strNos & "'", "'" & strInvoice & "'") = False Then Exit Sub
                    On Error GoTo errH
                End If
                'Ʊ�ݽӿ�
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
        Case 2  '�˷�
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
            Else
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
    End Select
    '3.�������ʹ�õ�����ID
    lngLastUseID = mlng����ID
    Exit Sub
ClearInvoice:
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(Split(strClearNOs, ","))
            strPrintNO = Split(strClearNOs, ",")(i)
            strSQL = "Zl_Ʊ����ʼ��_Update('" & strPrintNO & "','',4)"
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
    'Լ��[0��0���]ΪƱ�ݽ��
    If Sheet Is Nothing Then Exit Sub
    If Sheet.Cols = 0 Then Exit Sub
    If Sheet.ColWidth(0) <> 0 Then Exit Sub
    
    With mPrintSheet
        If .blnCalcMoney = False Then Exit Sub
        
        If .lngPrePage <> Page Then
            If .lngPrePage > 0 And .blnUseOnlyOneInvoice = False Then
                '��ǰҳ�ű仯���Ҳ��Ǵ�ӡֵʹ��һ�ŷ�Ʊ���򱣴���һҳ������
                Call UpdateInvoiceMoney(.arrInvoice(.lngPrePage - 1), .dblInvoiceMoney)
                .dblInvoiceMoney = 0
            ElseIf .lngPrePage = 0 Then
                .dblInvoiceMoney = 0
            End If
            
            .lngPrePage = Page
            .lngGridCount = 0
        End If
        
        '���ж�����ʱ���Ե�һ�����Ϊ׼
        If Row = 1 Then .lngGridCount = .lngGridCount + 1
        If .lngGridCount > 1 Then Exit Sub
        
        '�ۼƽ��
        .dblInvoiceMoney = .dblInvoiceMoney + Val(Sheet.TextMatrix(.lngCurPrintRow, 0))
        
        '�ۼƱ���кţ��������ǲ�����ҳ����
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
    '����Ʊ�ݽ��
    Dim strSQL As String
    
    On Error GoTo errHandle
    'Zl_Ʊ��ʹ����ϸ_���½��
    strSQL = "Zl_Ʊ��ʹ����ϸ_���½��("
    '  ����id_In   Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & mlng����ID & ","
    '  ��Ʊ��_In   Ʊ��ʹ����ϸ.����%Type,
    strSQL = strSQL & "'" & strInvoice & "',"
    '  Ʊ�ݽ��_In Ʊ��ʹ����ϸ.Ʊ�ݽ��%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  Ʊ��_In     Ʊ��ʹ����ϸ.Ʊ��%Type := 1
    strSQL = strSQL & "" & IIf(gblnSharedInvoice, 1, 4) & ")"
    zlDatabase.ExecuteProcedure strSQL, "����Ʊ�ݽ��"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
