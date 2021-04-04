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
Private mobjFactProperty As clsFactProperty
Private mobjInvoice As clsInvoice
Private mbytInFun As Byte               '1-�µ���ӡ,2-�ش�,3-��Ʊ��ӡ
Private mlng����ID As Long              '�ϴ�����ID
Private mstrPrintNO As String           '���ʵ��ݺ�
Private mlngBalanceID As Long           '����ID
Private mstrInvoice As String           '��ʼƱ�ݺ�
Private mdateBalance As Date            '���ʻ��ش��ʱ��
Private mblnPrinted As Boolean          '��ӡƱ�����������Ƿ�ɹ�
Private mblnInitInvoice As Boolean

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

 

Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
    Set mobjFactProperty = Nothing
    Set mobjInvoice = Nothing
    
    mbytInFun = 0
    mlng����ID = 0
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
                '�������һҳ������
                Call UpdateInvoiceMoney(.arrInvoice(.lngPrePage - 1), .dblInvoiceMoney)
            End If
        End If
    End With
End Sub

Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrInvoice As Variant)
    Dim strSQL As String, i As Integer, strInvoices As String
    Dim cllPro As Collection
    Dim strUserType As String, bytKind As Byte '0:סԺҽ�Ʒ��վ�,1-����ҽ�Ʒ��վ�
    
    With mPrintSheet
        .blnCalcMoney = True
        .lngPrePage = 0
        .lngCurPrintRow = 0
        .blnUseOnlyOneInvoice = False
    End With
    
    'û��Ʊ�ݺ�,�ϸ����Ʊ��ʱ����ӡ,���ϸ����Ʊ��ʱֻ��ӡ������Ʊ������
    If mblnInitInvoice = False Then
        mobjInvoice.zlInitCommon glngSys, gcnOracle, gstrDBUser
        mblnInitInvoice = True
    End If
    If mstrInvoice = "" Then
        Cancel = mobjFactProperty.�ϸ����
        mblnPrinted = Not mobjFactProperty.�ϸ����
        mPrintSheet.blnCalcMoney = False '������Ʊ�ݽ��
        Exit Sub
    End If
    Set cllPro = New Collection
    strUserType = ""
    If mobjFactProperty.ʹ����� <> "" Then strUserType = "(" & mobjFactProperty.ʹ����� & ")"
    mblnPrinted = False
    '1.�ϸ����Ʊ��ʱ������ʵ�ʵ�Ʊ������,���¼������ID��Ʊ�ݺ�
    If mobjFactProperty.�ϸ���� Then
        If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.����, mobjFactProperty.Ʊ��, _
            mobjFactProperty.ʹ�����, mlng����ID, mobjFactProperty.��������ID, mlng����ID, TotalPages, mstrInvoice) = False Then
            Cancel = True: Exit Sub
        End If
       If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case -1
                    MsgBox IIf(mbytInFun = 3, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "��Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "��û���㹻�����ú͹��õ�Ʊ��" & strUserType & ",������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox IIf(mbytInFun = 3, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "��Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "��û���㹻�ĵĹ���Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -3
                    MsgBox IIf(mbytInFun = 3, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "��Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & mstrInvoice & "]���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ�" & _
                        "������������Ч��Ʊ�ݺź��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -4
                    MsgBox IIf(mbytInFun = 3, "�����˷ѷ�Ʊ(��Ʊ)��ӡ", "����[" & mstrPrintNO & "]") & "��Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & mstrInvoice & "]���ڵ���������û���㹻��Ʊ�ݣ�" & _
                        "���ȴ�ӡ����Ʊ��,���굱ǰ�������κ�,�ش�õ��ݣ�", vbInformation, gstrSysName
                Case Else
                    MsgBox "Ʊ��������Ϣ����ʧ�ܣ������������" & IIf(mbytInFun = 3, "�ش�õ��ݣ�", "�ش򵥾�[" & mstrPrintNO & "]��"), vbInformation, gstrSysName
            End Select
            Cancel = True: Exit Sub
        End If
    End If
    
    '2.����Ʊ��ʹ������
    bytKind = IIf(mobjFactProperty.Ʊ�� = 3, 0, 1)
    On Error GoTo errH
    Select Case mbytInFun
        Case 1
            strSQL = "zl_���˽���Ʊ��_Insert('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng����ID) & _
                ",'" & UserInfo.���� & "',To_Date('" & Format(mdateBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & TotalPages & "," & bytKind & ")"
        
        Case 2
            strSQL = "zl_���˽��ʼ�¼_RePrint('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng����ID) & _
                ",'" & UserInfo.���� & "'," & TotalPages & "," & bytKind & ")"
        Case 3 '��Ʊ��ӡ
            'Zl_���˽��ʼ�¼_Reprint
            strSQL = "Zl_���˽��ʼ�¼_Reprint("
            '  No_In       ����Ԥ����¼.No%Type,
            strSQL = strSQL & "'" & mstrPrintNO & "',"
            '  Ʊ�ݺ�_In   Ʊ��ʹ����ϸ.����%Type,
            strSQL = strSQL & "'" & mstrInvoice & "',"
            '  ����id_In   Ʊ��ʹ����ϸ.����id%Type,
            strSQL = strSQL & "" & ZVal(mlng����ID) & ","
            '  ʹ����_In   Ʊ��ʹ����ϸ.ʹ����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  Ʊ������_In Number,
            strSQL = strSQL & "" & TotalPages & ","
            '  Ʊ��_In     Number := 0, --0:סԺҽ�Ʒ��վ�,1-����ҽ�Ʒ��վ�
            strSQL = strSQL & bytKind & ","
            '  ��Ʊ��ӡ_In Number := 0, --0:�����ش�,1-����ʱ���Ʊ��ӡ
            strSQL = strSQL & "" & 1 & ","
            '  ʹ��ʱ��_In Date:=Null
            strSQL = strSQL & "" & "To_date('" & Format(mdateBalance, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')" & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSQL, "Ʊ����������")
    mblnPrinted = True
    
    '3.�������õ�Ʊ�ݺ���Ϣ
    For i = 1 To TotalPages
        strInvoices = strInvoices & "," & mstrInvoice
        If i < TotalPages Then mstrInvoice = zlCommFun.IncStr(mstrInvoice)
    Next
    strInvoices = Mid(strInvoices, 2)
    If strInvoices <> "" Then arrInvoice = Split(strInvoices, ",")
    
    mPrintSheet.arrInvoice = arrInvoice
    
    '���ϸ����Ʊ��ʱ���浽ע���
    If Not mobjFactProperty.�ϸ���� Then
        zlDatabase.SetPara "��ǰ����Ʊ�ݺ�", mstrInvoice, glngSys, 1137
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
                        Optional str�ɿ� As String, Optional str�Ҳ� As String, Optional lngPatientID As Long, _
                        Optional intLocalFormat As Integer, Optional blnPrintBillEmpty As Boolean = False, _
                        Optional blnInsurePrint As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݴ�ӡ
    '���:bytInfun:1-�µ���ӡ,2-�ش�,3-��Ʊ��ӡ
    '       strNO:���ʵ��ݺ�,��������
    '       lngBalanceID:����ID
    '       objFactProperty-��Ʊ���Կ���
    '       lngLastUseID:���ʹ�õ���������ID,����ʱΪ0
    '       lngShareUseID:��������
    '       strUseType:ʹ�����
    '       strInvoice:��ʼƱ�ݺţ���������,���ϸ����Ʊ��ʱ�������,�ϸ����ʱ����ǰ��ǰ��鲻��Ϊ��
    '       dateBalance :����ʱ��,���µ���ӡ�Ŵ���
    '       lngPatientID:��Լ��λ���ʰ����˷ֱ��ӡ,ÿ�δ�ӡ���뵱ǰ����ID
    '       intLocalFormat:��ָ���ĸ�ʽ��ӡ
    '       blnInsurePrint:�Ƿ�ҽ���ӿڴ�ӡ
    '����:
    '       blnPrintBillEmpty-�Ƿ��ӡ��Ʊ��(55052)
    '����:
    '����:���˺�
    '����:2011-05-03 17:44:08
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
    '1.��������
    mbytInFun = bytInfun: mstrPrintNO = strNO
    mlngBalanceID = lngBalanceID: mlng����ID = objFactProperty.LastUseID
    mstrInvoice = strInvoice: mdateBalance = dateBalance
    Set mobjFactProperty = objFactProperty
    
    If objFactProperty.Ʊ�� = 3 Then
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
    'ѡ��Ĵ�ӡ��ʽ
    strFormat = IIf(intLocalFormat <= 0, "", "ReportFormat=" & intLocalFormat)
    mblnPrinted = False
    
    '2.��ӡ����
    Select Case mbytInFun
        Case 1  '�µ���ӡ
            If Not gobjTax Is Nothing And gblnTax Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)   '���ô�ӡ����������ӡ��ֻ������Ʊ��ʹ������
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
                    Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)   '���ô�ӡ����������ӡ��ֻ������Ʊ��ʹ������
                Else
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "����ID=" & mlngBalanceID, "����ID=" & lngPatientID, "PrintEmpty=0", str�ɿ�, str�Ҳ�, strFormat, 2)
                    If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                    If Not mblnPrinted Then GoTo ClearInvoice
                End If
            End If
        Case 2  '�ش�
            If Not gobjTax Is Nothing And gblnTax Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)
                If IsArray(arrInvoice) Then
                    mstrInvoice = arrInvoice(0)
                Else
                    mstrInvoice = arrInvoice
                End If
                If Not mblnPrinted Then Exit Sub
                
                If Not gobjTax Is Nothing And gblnTax Then
                    MsgBox "����׼����֮��ȷ����ʼ��ӡ��", vbInformation, gstrSysName
                    gstrTax = gobjTax.zlTaxInReput(gcnOracle, mlngBalanceID)
                    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
                End If
            Else
                If gblnBillPrint Then
                    If gobjBillPrint.zlRePrintBill("", mlngBalanceID, strInvoice) = False Then Exit Sub
                End If
                
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "����ID=" & mlngBalanceID, "����ID=" & lngPatientID, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
        Case 3 '��Ʊ��ӡ
            If Not gobjTax Is Nothing And gblnTax Then
                Call mobjReport_BeforePrint(strReportNO, 1, True, arrInvoice)
                If IsArray(arrInvoice) Then
                    mstrInvoice = arrInvoice(0)
                Else
                    mstrInvoice = arrInvoice
                End If
                If Not mblnPrinted Then Exit Sub
                
                If Not gobjTax Is Nothing And gblnTax Then
                    MsgBox "����׼����֮��ȷ����ʼ��ӡ��", vbInformation, gstrSysName
                    gstrTax = gobjTax.zlTaxInReput(gcnOracle, mlngBalanceID)
                    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
                End If
            Else
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "����ID=" & mlngBalanceID, "����ID=" & lngPatientID, "PrintEmpty=0", "", "", strFormat, 2)
                If blnPrintBillEmpty = False Then blnPrintBillEmpty = mobjReport.DataIsEmpty    '55052
                If Not mblnPrinted Then Exit Sub
            End If
    End Select
    '3.�������ʹ�õ�����ID
    mobjFactProperty.LastUseID = mlng����ID
    Exit Sub
    
ClearInvoice:
    On Error GoTo errH
    strSQL = "Zl_Ʊ����ʼ��_Update('" & strNO & "','',3)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errH:
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
    strSQL = strSQL & "" & mobjFactProperty.Ʊ�� & ")"
    zlDatabase.ExecuteProcedure strSQL, "����Ʊ�ݽ��"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
