VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "Ʊ�ݴ�ӡ"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mstrPrintNO As String           'Ҫ��ӡ�ĵ��ݺ�
Private mstrInvoice As String           'Ҫ��ӡ��Ʊ�ݺ�
Private mEditType As gCardType          '��������
Private mlng����ID As Long              '��������ID
Private mstrUseType As String           'ʹ�����
Private mdtPrintdate As Date            '��ӡʱ��
Private mUserName As String             'ʹ����
                                
Public Sub PrintBill(ByVal strNO As String, ByVal strCardNo As String, _
                     ByVal strInvoice As String, ByVal lngCardTypeID As Long, ByVal blnPrint As Boolean, _
                     ByVal EditType As gCardType, ByVal bytPrintFormat As Byte, ByVal lng����ID As Long, _
                     ByVal strUseType As String, ByVal dtPrintdate As Date, ByVal UserName As String, _
                     Optional blnPrepayPrint As Boolean = False, Optional strPrePayNo As String = "", _
                     Optional lngԤ������ID As Long = 0, Optional datԤ��ʱ�� As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�з���Ʊ�ݴ�ӡ
    '������strNO           ���ݺ�
    '      strPrePayNo     Ԥ������
    '      strCardNo       ����
    '      lngԤ������ID   Ԥ������ID
    '      strInvoice      Ʊ�ݺ�
    '      lngCardTypeID   �����ID
    '      blnPrint        �Ƿ��ӡ
    '      blnPrepayPrint  �Ƿ��ӡԤ����
    '      EditType        ��������
    '      bytPrintFormat  ��ӡ��ʽ:����|�󶨿�
    '      lng����ID       ��������ID
    '      strUseType      ʹ�����
    '      dtPrintdate     ��ӡʱ��
    '      UserName        ʹ����
    '      blnPrepayPrint  �Ƿ��ӡԤ����
    '      strPrePayNo     Ԥ������
    '����:���ϴ�
    '����:2014-04-10 13:41:24
    '�����:57950
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFormat As String
    On Error GoTo Errhand
    mstrPrintNO = strNO
    mstrInvoice = strInvoice
    mlng����ID = lng����ID
    mstrUseType = strUseType
    mdtPrintdate = dtPrintdate
    mUserName = UserName
    
    If blnPrepayPrint Then
        '��ӡԤ��Ʊ��
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & strPrePayNo, "����ID=" & lngԤ������ID, "�տ�ʱ��=" & Format(datԤ��ʱ��, "yyyy-mm-dd HH:MM:SS"), 2)
    End If
    
    If Not blnPrint Then Exit Sub
    strFormat = IIf(bytPrintFormat = 0, "", "ReportFormat=" & bytPrintFormat)
    
    If mEditType = Cr_�󶨿� Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1107", Me, "�����ID=" & lngCardTypeID, "NO=" & strCardNo, "����=" & strCardNo, "PrintEmpty=0", strFormat, 2)
    ElseIf gbln�շѷ�Ʊ Then
        Set mobjReport = New clsReport
        Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1107", Me, "�����ID=" & lngCardTypeID, "NO=" & strNO, "����=" & strCardNo, "PrintEmpty=0", strFormat, 2)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1107", Me, "�����ID=" & lngCardTypeID, "NO=" & strNO, "����=" & strCardNo, "PrintEmpty=0", strFormat, 2)
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
    Dim lng����ID As Long
    Dim strSQL As String
    arrBill = Split(mstrInvoice, ",")
    On Error GoTo errH
    If gblnBill���� Then
        lng����ID = GetInvoiceGroupID(1, TotalPages, mlng����ID, glngShareUseID, mstrInvoice, mstrUseType)
        If lng����ID <= 0 Then
            Select Case lng����ID
                Case -1
                    MsgBox "����[" & mstrPrintNO & "]" & "����Ҫ" & TotalPages & "��Ʊ�ݣ�" & vbCrLf & _
                        "��û���㹻�����ú͹��õ�Ʊ�ݣ�������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "����[" & mstrPrintNO & "]" & "����Ҫ" & TotalPages & "��Ʊ�ݣ�" & vbCrLf & _
                        "��û���㹻�ĵĹ���Ʊ�ݣ�������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -3
                    MsgBox "����[" & mstrPrintNO & "]" & "����Ҫ" & TotalPages & "��Ʊ�ݣ�" & vbCrLf & _
                        "Ʊ�ݺ�[" & mstrInvoice & "]���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ�" & _
                        "������������Ч��Ʊ�ݺź��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -4
                    MsgBox "����[" & mstrPrintNO & "]" & "����Ҫ" & TotalPages & "��Ʊ�ݣ�" & vbCrLf & _
                        "Ʊ�ݺ�[" & mstrInvoice & "]���ڵ���������û���㹻��Ʊ�ݣ�" & _
                        "���ȴ�ӡ����Ʊ��,���굱ǰ�������κ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case Else
                    MsgBox "Ʊ��������Ϣ����ʧ�ܣ�������������ش򵥾�[" & mstrPrintNO & "]��", vbInformation, gstrSysName
            End Select
            Cancel = True: Exit Sub
        End If
    End If
    strSQL = "Zl_���˷���Ʊ��_Print("
    '  No_In           Varchar2,
    strSQL = strSQL & "'" & Replace(mstrPrintNO, "'", "") & "'" & ","
    '  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
    strSQL = strSQL & "'" & mstrInvoice & "',"
    '  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & ZVal(lng����ID) & ","
    '  ʹ����_In       Ʊ��ʹ����ϸ.ʹ����%Type,
    strSQL = strSQL & "'" & mUserName & "',"
    '  ʹ��ʱ��_In     Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
    strSQL = strSQL & "To_Date('" & Format(mdtPrintdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  ��������_In     Number
    strSQL = strSQL & IIf(mEditType = Cr_����, 5, 1) & ","
    '  Ʊ������_In     Number := 1,
    strSQL = strSQL & "" & TotalPages & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "Ʊ����������")
    
    '���ϸ����Ʊ��ʱ���浽ע���
    '���±���Ʊ��
    If Not gblnBill���� Then
        zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", mstrInvoice, glngSys, 1121
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
    '����:�ش�Ʊ��(��ģʽ)
    Dim strFormat As String
    On Error GoTo errH
    mstrInvoice = ""
  
    If strSelect = "����" Then
        If strCardNo = "" Then ShowMsgbox "ûѡ����ص�ҽ�ƿ�": Exit Sub
        strFormat = IIf(bytPrintPayCard = 0, "", "ReportFormat=" & bytPrintPayCard)
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1107", Me, "�����ID=" & lngCardTypeID, "NO=" & strCardNo, "����=" & strCardNo, "�ɿ�=" & 0, "�Ҳ�=" & 0, "PrintEmpty=0", strFormat, 2)
    Else
        'Ԥ����ӡ
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
                             ByVal bytPrintPayCard As Byte, Optional ByVal bln�ش� As Boolean) As Boolean
    '--------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ǰ�տ��¼���´�ӡһ��Ʊ��(��ģʽ)
    '���:   strUseType-ʹ�����

    '        blnVirtualPrint-ҽ���ӿ��ڵ��ô�ӡ��HISֻ��Ʊ�Ų�ʵ�ʴ�ӡ
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-11-19 17:18:19
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    Dim strInvoice As String
    Dim blnValid As Boolean, blnInput As Boolean
    Dim lng����ID As Long, strBackInvoice As String
    Dim blnReprint As Boolean, strFormat As String
    
    On Error GoTo errH
    '����ϸ����Ʊ��ʹ��
    If gblnBill���� Then
        If bln�ش� Then
            lng����ID = CheckUsedBill(1, glngShareUseID, , strUseType)
            Select Case lng����ID
                Case -1
                    MsgBox "��û�����ú͹��õĹҺ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End Select
            If lng����ID <= 0 Then Exit Function
        End If
        If intPrintMode = 3 Then
            '��ȡ�ջ�Ʊ��
            strSQL = _
            "   Select A.����" & vbNewLine & _
            "   From Ʊ��ʹ����ϸ A" & vbNewLine & _
            "   Where A.���� = 1 And a.ԭ�� <> 6 " & vbNewLine & _
            "       And A.Ʊ�� = 1 And A.��ӡid = (Select Max(ID) From Ʊ�ݴ�ӡ���� Where �������� = [2] And NO = [1])" & vbNewLine & _
            "   Order By ����"
            Set rsInvoice = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ջ�Ʊ��", strPrintNo, 5)
            Do While Not rsInvoice.EOF
                strBackInvoice = strBackInvoice & "," & rsInvoice!����
                rsInvoice.MoveNext
            Loop
            If strBackInvoice <> "" Then strBackInvoice = Mid(strBackInvoice, 2)
        End If
        blnReprint = bln�ش�
    End If
    
     'ȡ��һ��Ʊ�ݺ���
    If Not gblnBill���� Then
        '�п����ǵ�һ��ʹ��
        Do
            blnInput = False
            '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
            strInvoice = zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, 1121)
            mstrInvoice = strInvoice
            
            If strInvoice = "" Then
                strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                "", frmParent.Left + 1500, frmParent.Top + 1500))
                blnInput = True
            Else
                strInvoice = zlCommFun.IncStr(strInvoice)
                strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                blnInput = True
            End If
                
            '�û�ȡ������,�����ӡ
            If strInvoice = "" Then
                If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnValid = True
            Else
                '���������Ч��
                If blnInput Then
                    If zlCommFun.ActualLen(strInvoice) <> gbyt�շ� Then
                        MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & gbyt�շ� & " λ��", vbInformation, gstrSysName
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
                '����Ʊ�����ö�ȡ
                blnInput = False
                strInvoice = GetNextBill(lng����ID)
                If strInvoice = "" Then
                    '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
                    strInvoice = UCase(InputBox("�޷�����Ʊ�����������ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                    vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If
                
                '�û�ȡ������,����ӡ
                If strInvoice = "" Then Exit Function
                
                '���������Ч��
                If blnInput Then
                    If GetInvoiceGroupID(1, 1, lng����ID, glngShareUseID, strInvoice, strUseType) = -3 Then
                        MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
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
    
    mlng����ID = lng����ID
    mstrInvoice = strInvoice
    'ִ�����ݴ���
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    strFormat = IIf(bytPrintPayCard = 0, "", "ReportFormat=" & bytPrintPayCard)
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1107", Me, "�����ID=" & lngCardTypeID, "NO=" & strPrintNo, "����=" & strCardNo, "PrintEmpty=0", strFormat, 2)
    
    RePrintBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


