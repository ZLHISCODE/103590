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
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "frmReplenishTheBalancePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------------------------------------
'���������ر���
Private mbytInFun As Byte                 '1-�µ���ӡ,2-�ش�,3-�˷Ѵ�ӡ; 4-����Ʊ��;6-�˷�Ʊ��(��Ʊ)��ӡ
Private mobjFactProperty As clsFactProperty
Private mintInsure As Integer
Private mstrReclaimInvoice As String    'Ҫ����յķ�Ʊ��,��1-����ϵͳԤ���������Ʊ�ź�2-�����û������������Ʊ����Ч
'--------------------------------------------------------------------------------------
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mlng����ID As Long              '�ϴ�����ID
Private mstrPrintNO As String           'Ҫ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...
Private mstrInvoice As String           '��ʼƱ�ݺ�
Private mdatFeeDate As Date             '���õ������ݵĵǼ�ʱ��
Private mblnPrinted As Boolean          'Ʊ�����������Ƿ�ɹ�(�Ƿ��Ѵ�ӡ)
Private mstrPrivs As String
Private mstrUseType As String
Private mbln����Ʊ�� As Boolean
Private mobjInvoice As clsInvoice
Private mblnֻ��һ��Ʊ�� As Boolean
Private mlngModule As Long

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
 

Public Sub ReportPrint(ByVal bytInFun As Byte, ByVal strNos As String, ByVal intInsure As Integer, _
                        ByVal objFactProperty As clsFactProperty, _
                        ByVal strReclaimInvoice As String, _
                        ByRef lngLastUseID As Long, ByVal strInvoice As String, ByVal datFeeDate As Date, _
                        Optional blnVirtualPrint As Boolean, _
                        Optional ByVal blnDelRecord As Boolean, _
                        Optional blnPrintBillEmpty As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ʊ�ݴ�ӡ,�������
    '���:bytInfun :1-�µ���ӡ,2-�ش�,3-�˷Ѵ�ӡ,4-����Ʊ��(ֻ��:2-��ϵͳԤ�������3-�û��Զ�����ʱ��ת��),6-�˷�Ʊ��(��Ʊ)��ӡ
    '       strNOs - �µ�ʱҪ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...,
    '                   - �˷�Ʊ��(��Ʊ)��ӡʱ������������
    '       strReclaimInvoice-Ҫ����յķ�Ʊ��,����ö��ŷ���'F0000001','F0000002',...
    '       lngLastUseID-���ʹ�õ���������ID,����ʱΪ0
    '       strInvoice-��ʼƱ�ݺţ���������,���ϸ����Ʊ��ʱ�������,�ϸ����ʱ����ǰ��ǰ��鲻��Ϊ��
    '       datFeeDate-���ý���ʱ��
    '       blnVirtualPrint-ҽ���ӿ��ڵ��ô�ӡ��HISֻ��Ʊ�Ų�ʵ�ʴ�ӡ
    '       blnDelRecord-�ش�ʱ���Ƿ��Ƕ��˷Ѽ�¼�����ش�(Ŀǰֻ�б���ҽ��(ҽ���ӿڴ�ӡƱ��)������)
    '       lngShareUseID-��������
    '       strUseType-ʹ�����
    ' ����:
    '   blnPrintBillEmpty-�Ƿ��ӡ�Ŀձ�����()
    '����:���˺�
    '����:2014-09-24 18:11:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, strPrintNO As String, blnPrint As Boolean, blnTrans As Boolean
    Dim strReportNO As String, strSQL As String, strClearNOs As String, strFormat As String, lngBalanceID As Long
    Dim blnNotPrint As Boolean, varTmp As Variant '�ձ�������Ҫ��Ϊ�˴�����ú����ķ���ֵ
    Dim str��Ʊ�� As String, intƱ������ As Integer
    blnPrintBillEmpty = False
    mbln����Ʊ�� = False
    
    mbytInFun = bytInFun: mdatFeeDate = datFeeDate: mlngModule = 1124
    
    
    mlng����ID = lngLastUseID: mstrInvoice = strInvoice: mstrReclaimInvoice = strReclaimInvoice
    Set mobjFactProperty = objFactProperty: mintInsure = intInsure
    If bytInFun <> 6 Then strNos = IIf(InStr(1, strNos, "'") = 0, "'" & Replace(strNos, ",", "','") & "'", strNos)
    
    Me.Caption = "��ӡ"
    
    '1.��������
    If mbytInFun = 6 Then '�˷�Ʊ��(��Ʊ)��ӡ
        strReportNO = "ZL" & glngSys \ 100 & "_BILL_1124_3"
    Else
        strReportNO = "ZL" & glngSys \ 100 & "_BILL_1124"
    End If
    strFormat = IIf(objFactProperty.��ӡ��ʽ = 0, "", "ReportFormat=" & objFactProperty.��ӡ��ʽ)
    
    mstrPrintNO = "": mblnPrinted = False
    blnNotPrint = (Not gobjTax Is Nothing And gblnTax) Or blnVirtualPrint
    
    '2.��ӡ����
    Select Case mbytInFun
        Case 1 '�µ���ӡ���ش�Ʊ��
             
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp) '���ô�ӡ����������ӡ��ֻ������Ʊ��ʹ������
                If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice '�޸�ʱ,ֻ����µ��ݵĿ�ʼƱ�ݺ�
                Call TaxInterface(1, mstrPrintNO, "")
            Else
               'Ʊ�ݽӿ�
                If BillPrint(1, mstrPrintNO, "", "", strClearNOs) = False Then: GoTo ClearInvoice
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty
                If Not mblnPrinted Then strClearNOs = Replace(strNos, "'", ""): GoTo ClearInvoice
            End If
 
        Case 2, 4 '�ش�
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
                Call TaxInterface(2, mstrPrintNO, "")       '��ӡ˰��Ʊ��
                ''����ҽ���ش�ӿ�
                If InsureReprint(blnVirtualPrint, Replace(Split(strNos, ",")(0), "'", ""), lngBalanceID, blnDelRecord, strInvoice) = False Then Exit Sub
            Else
                'Ʊ�ݽӿ�
                 If BillPrint(2, mstrPrintNO, "", strInvoice, strClearNOs) = False Then Exit Sub
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
        Case 3  '�˷�
        
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
                Call TaxInterface(3, mstrPrintNO, "")
            Else
                If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit Sub
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "��Ʊ��=FactNO", "NO=" & mstrPrintNO, "PrintEmpty=0", "", "", strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
        Case 6 '��Ʊ��ӡ
            mstrPrintNO = strNos
            If blnNotPrint Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, varTmp)
                If Not mblnPrinted Then Exit Sub
'                Call TaxInterface(3, mstrPrintNO, "")
            Else
'                If BillPrint(3, mstrPrintNO, "", strInvoice, "") = False Then Exit Sub
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "�������=" & Val(strNos), "PrintEmpty=0", strFormat, 2)
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
            strSQL = "Zl_Ʊ����ʼ��_Update('" & strPrintNO & "','',1)"
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
    
    If mblnֻ��һ��Ʊ�� Then
        mPrintSheet.blnUseOnlyOneInvoice = True
        TotalPages = 1 '�շ�ÿ�δ�ӡֻ��һ��Ʊ��
    End If
    
    'û��Ʊ�ݺ�,�ϸ����Ʊ��ʱ����ӡ,���ϸ����Ʊ��ʱֻ��ӡ������Ʊ������
    If mstrInvoice = "" Then
        Cancel = mobjFactProperty.�ϸ����
        mblnPrinted = Not mobjFactProperty.�ϸ����
        mPrintSheet.blnCalcMoney = False '������Ʊ�ݽ��
        Exit Sub
    End If
    
    
    If CheckInvoiceValied(TotalPages, mbytInFun = 6) = False Then Cancel = True: Exit Sub
    
    On Error GoTo errH
    
    '2.����Ʊ������
    Select Case mbytInFun
        Case 1
            strSQL = "Zl_�������Ʊ��_Insert('" & Replace(mstrPrintNO, "'", "") & "','" & mstrInvoice & "'," & ZVal(mlng����ID) & ",'" & UserInfo.���� & "'," & _
                     "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),0," & TotalPages & ")"
        Case 2, 3
            '����Ƕ��ţ�ֻ��Ҫ��һ�ŵ��ݺž�����(�޸Ķ����е�һ��ʱ,���һ�����µ�)
            strSQL = "Zl_�������Ʊ��_Reprint('" & Replace(Split(mstrPrintNO, ",")(0), "'", "") & "','" & mstrInvoice & "'," & ZVal(mlng����ID) & ",'" & UserInfo.���� & "'," & _
                    "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & IIf(mbytInFun = 2, "0", "1") & "," & TotalPages & ")"
        Case 6 '�˷ѷ�Ʊ(��Ʊ)
            'Zl_��������˷�Ʊ��_Insert
            strSQL = "Zl_��������˷�Ʊ��_Insert("
            '  �������_In   ����Ԥ����¼.�������%Type,
            strSQL = strSQL & "" & Val(mstrPrintNO) & ","
            '  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
            strSQL = strSQL & "'" & mstrInvoice & "',"
            '  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
            strSQL = strSQL & "" & ZVal(mlng����ID) & ","
            '  ʹ����_In       Ʊ��ʹ����ϸ.ʹ����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ʹ��ʱ��_In     Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
            strSQL = strSQL & "To_Date('" & Format(mdatFeeDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  Ʊ������_In Number:=1
            strSQL = strSQL & "" & TotalPages & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSQL, "Ʊ����������")
    mblnPrinted = True
    
    '3.�������õ�Ʊ�ݺ���Ϣ
    For i = 1 To TotalPages
        strInvoices = strInvoices & "," & mstrInvoice
        If i < TotalPages Then mstrInvoice = zlStr.Increase(mstrInvoice)
    Next
    strInvoices = Mid(strInvoices, 2)
    arrInvoice = Split(strInvoices, ",")
    
    mPrintSheet.arrInvoice = arrInvoice
        
    '���ϸ����Ʊ��ʱ���浽ע���
    If Not mobjFactProperty.�ϸ���� Then
        zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", mstrInvoice, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub
Private Function CheckInvoiceValied(Optional int���� As Integer = 1, _
    Optional ByVal blnDelFeePrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鷢Ʊ�Ƿ�Ϸ�(�ϸ����Ʊ��ʱ)
    '���:int���� -��Ҫ�ķ�Ʊ����
    '   blnDelFeePrint-�˷ѷ�Ʊ(��Ʊ)��ӡ
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-09-24 17:49:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjFactProperty.�ϸ���� Then CheckInvoiceValied = True: Exit Function
    
    '1.�ϸ����Ʊ��ʱ������ʵ�ʵ�Ʊ������,���¼������ID��Ʊ�ݺ�Property.ʹ�����, mlng����
    If mobjInvoice.zlGetInvoiceGroupID(mlngModule, UserInfo.����, EM_�շ��վ�, mobjFactProperty.ʹ�����, mobjFactProperty.��������ID, mlng����ID, mlng����ID, int����, mstrInvoice) = False Then Exit Function
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


Private Sub TaxInterface(ByVal byt���� As Byte, ByVal strPrintNO As String, ByVal strModiNos As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˰�ش�ӡ�ӿ�
    '���:byt����-1-������ӡ(���޸�);2-�ش�;3-�˷�
    '        strPrintNO-Ҫ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...
    '        strModiNos-�޸Ķ൥���е�һ��ʱ,ָ�ö��ŵ��ݵ�����NO���ö��ŷָ�:'F0000001','F0000002',...
    '����:���˺�
    '����:2013-03-27 14:24:03
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    'δ����˰��,ֱ�ӷ���
    If Not gblnTax Then Exit Sub
    If byt���� = 3 Then
        '�˷�
        gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strPrintNO)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        gstrTax = gobjTax.zlTaxOutReput(gcnOracle, strPrintNO)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If byt���� = 2 Then
        '�ش�
        MsgBox "����׼����֮��ȷ����ʼ��ӡ��", vbInformation, gstrSysName
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
Private Function BillPrint(ByVal byt���� As Byte, ByVal strPrintNO As String, _
    ByVal strModiNos As String, ByRef strInvoice As String, ByRef strClearNOs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݴ�ӡ�ӿ�
    '���:byt����-1-������ӡ(���޸Ĵ�ӡ);2-�ش��ӡ;3-�˷�
    '        strPrintNO-Ҫ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...
    '        strModiNos-�޸Ķ൥���е�һ��ʱ,ָ�ö��ŵ��ݵ�����NO���ö��ŷָ�:'F0000001','F0000002',...
    '         strInvoice-��Ʊ��(�ش�ʱ��Ч)
    '����:strClearNOs-��Ҫ����ĵ��ݺ�
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-03-27 14:36:28
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gblnBillPrint Then BillPrint = True: Exit Function
    If byt���� = 3 Then
        '�˷�
        '�˷�����֮ǰ�ȵ���Ʊ���ջأ�zlEraseBill
        BillPrint = gobjBillPrint.zlRePrintBill(strPrintNO, 0, strInvoice)
        Exit Function
    End If
    If byt���� = 2 Then
        '�ش�
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
    '����:2014-09-24 18:02:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    If Not blnVirtualPrint Then InsureReprint = True: Exit Function
    '81222
    If lng����ID = 0 Then
        strSQL = "Select Max(����ID) As ����ID From ���ò����¼ Where NO= [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
        If Not rsTmp.EOF Then
            lng����ID = rsTmp!����ID
        End If
    End If
    Call gclsInsure.RePrintBill(mintInsure, lng����ID, strInvoice)
    InsureReprint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mobjReport_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long)
    
    On Error GoTo errHandle
    'Լ��[0��0���]ΪƱ�ݽ��
    If Sheet Is Nothing Then Exit Sub
    If Sheet.COLS = 0 Then Exit Sub
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
    strSQL = strSQL & "" & 1 & ")"
    zlDatabase.ExecuteProcedure strSQL, "����Ʊ�ݽ��"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
