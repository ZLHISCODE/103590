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

Private mbytInFun As Byte               '1-�µ���ӡ,2-�ش�
Private mlng����ID As Long              '�ϴ�����ID
Private mstrPrintNO As String           '���ʵ��ݺ�
Private mlngBalanceID As Long           '����ID
Private mstrInvoice As String           '��ʼƱ�ݺ�
Private mdateBalance As Date            '���ʻ��ش��ʱ��
Private mblnPrinted As Boolean          '��ӡƱ�����������Ƿ�ɹ�
Private mbytKind As Byte

Private Sub Form_Load()
    Set mobjReport = New clsReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
    mbytInFun = 0
    mlng����ID = 0
    mstrPrintNO = ""
    mlngBalanceID = 0
    mstrInvoice = ""
    mdateBalance = CDate(0)
    mblnPrinted = False
End Sub


Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrInvoice As Variant)
    Dim strSQL As String, i As Integer, strInvoices As String
    
    'û��Ʊ�ݺ�,�ϸ����Ʊ��ʱ����ӡ,���ϸ����Ʊ��ʱֻ��ӡ������Ʊ������
    If mstrInvoice = "" Then
        Cancel = gblnStrictCtrl
        mblnPrinted = Not gblnStrictCtrl
        Exit Sub
    End If
    
    mblnPrinted = False
    '1.�ϸ����Ʊ��ʱ������ʵ�ʵ�Ʊ������,���¼������ID��Ʊ�ݺ�
    If gblnStrictCtrl Then
        mlng����ID = GetInvoiceGroupID(mbytKind, TotalPages, mlng����ID, glngShareUseID, mstrInvoice)
        If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case -1
                    MsgBox "����[" & mstrPrintNO & "]����Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "��û���㹻�����ú͹��õ�Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "����[" & mstrPrintNO & "]����Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "��û���㹻�ĵĹ���Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -3
                    MsgBox "����[" & mstrPrintNO & "]����Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & mstrInvoice & "]���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ�" & _
                        "������������Ч��Ʊ�ݺź��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -4
                    MsgBox "����[" & mstrPrintNO & "]����Ҫ" & TotalPages & "��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & mstrInvoice & "]���ڵ���������û���㹻��Ʊ�ݣ�" & _
                        "���ȴ�ӡ����Ʊ��,���굱ǰ�������κ�,�ش�õ��ݣ�", vbInformation, gstrSysName
                Case Else
                    MsgBox "Ʊ��������Ϣ����ʧ�ܣ�������������ش򵥾�[" & mstrPrintNO & "]", vbInformation, gstrSysName
            End Select
            Cancel = True: Exit Sub
        End If
    End If
    
    '2.����Ʊ��ʹ������
    On Error GoTo errH
    Select Case mbytInFun
        Case 1
            strSQL = "zl_���˽���Ʊ��_Insert('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng����ID) & _
                ",'" & UserInfo.���� & "',To_Date('" & Format(mdateBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & TotalPages & ")"
        
        Case 2
            strSQL = "zl_���˽��ʼ�¼_RePrint('" & mstrPrintNO & "','" & mstrInvoice & "'," & ZVal(mlng����ID) & _
                ",'" & UserInfo.���� & "'," & TotalPages & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSQL, "Ʊ����������")
    mblnPrinted = True
    
    '3.�������õ�Ʊ�ݺ���Ϣ
    For i = 1 To TotalPages
        strInvoices = strInvoices & "," & mstrInvoice
        mstrInvoice = IncStr(mstrInvoice)
    Next
    strInvoices = Mid(strInvoices, 2)
    arrInvoice = Split(strInvoices, ",")
        
    '���ϸ����Ʊ��ʱ���浽ע���
    If Not gblnStrictCtrl Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ǰ����Ʊ�ݺ�", mstrInvoice
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub

Public Sub ReportPrint(ByVal bytInfun As Byte, ByVal strNO As String, ByVal lngBalanceID As Long, _
                        ByRef lngLastUseID As Long, ByVal strInvoice As String, Optional ByVal dateBalance As Date, _
                        Optional str�ɿ� As String, Optional str�Ҳ� As String, Optional ByVal bytKind As Byte = 3)
'������ bytInfun        =   1-�µ���ӡ,2-�ش�
'       strNO           =   ���ʵ��ݺ�,��������
'       lngBalanceID    =   ����ID
'       lngLastUseID   <=>  ���ʹ�õ���������ID,����ʱΪ0
'       strInvoice      =   ��ʼƱ�ݺţ���������,���ϸ����Ʊ��ʱ�������,�ϸ����ʱ����ǰ��ǰ��鲻��Ϊ��
'       dateBalance     =   ����ʱ��,���µ���ӡ�Ŵ���
    Dim strReportNO As String, strSQL As String
    
    '1.��������
    mbytInFun = bytInfun
    mstrPrintNO = strNO
    mlngBalanceID = lngBalanceID
    mlng����ID = lngLastUseID
    mstrInvoice = strInvoice
    mdateBalance = dateBalance
    strReportNO = "ZL1_BILL_1862"
    mbytKind = bytKind
    
    '2.��ӡ����
    Select Case mbytInFun
        Case 1  '�µ���ӡ

            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "����ID=" & mlngBalanceID, str�ɿ�, str�Ҳ�, 2)
            If Not mblnPrinted Then GoTo ClearInvoice

        Case 2  '�ش�

            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "����ID=" & mlngBalanceID, "", "", 2)
            If Not mblnPrinted Then Exit Sub

    End Select
    
    '3.�������ʹ�õ�����ID
    lngLastUseID = mlng����ID
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
