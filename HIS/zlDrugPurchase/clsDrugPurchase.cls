VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub BillNumber(ByVal cnVal As ADODB.Connection, ByVal strPrivs As String, ByVal intListIndex As Integer)
'------------------------------
'����: �������͵��ŵ����������
'------------------------------
    Dim rs�汾��Ϣ As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    If cnVal Is Nothing Then
        MsgBox "����ZLHIS�����ݿ�����ʧ�ܣ���ֹ���룡", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    If gobjComLib Is Nothing Then
        MsgBox "����zl9ComLib����ʧ�ܣ���ֹ���룡", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    Set gcnOracle = cnVal

    Call gobjComLib.InitCommon(cnVal)              '�����ڲ�Connection
    
    gstrPrivs = strPrivs
    gintListIndex = intListIndex
    
    strSQL = "Select t.�汾�� From zlSystems T Where t.��� = 100"
    Set rs�汾��Ϣ = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ�汾��Ϣ")
    
    If rs�汾��Ϣ.RecordCount = 0 Then Exit Sub
    
    '���ֵ��õ��ĸ��汾����Ϊ�ò��������ְ汾
    If InStr(NVL(rs�汾��Ϣ!�汾��), "10.35.110") > 0 Then
        frmBillNumber35110.Show vbModal
    ElseIf InStr(NVL(rs�汾��Ϣ!�汾��), "10.34.10") > 0 Then
        frmBillNumber3410.Show vbModal
    ElseIf InStr(NVL(rs�汾��Ϣ!�汾��), "10.35") > 0 Then
        frmBillNumber.Show vbModal
    ElseIf InStr(NVL(rs�汾��Ϣ!�汾��), "10.34") > 0 Then
        frmBillNumber34.Show vbModal
    End If
    
    Exit Sub
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub
Public Sub DeliveryInvoice(ByVal cnVal As ADODB.Connection)
'------------------------------
'����: �����ͻ���Ʊ�����������
'------------------------------
    If cnVal Is Nothing Then
        MsgBox "����ZLHIS�����ݿ�����ʧ�ܣ���ֹ���룡", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    If gobjComLib Is Nothing Then
        MsgBox "����zl9ComLib����ʧ�ܣ���ֹ���룡", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    Set gcnOracle = cnVal

    Call gobjComLib.InitCommon(cnVal)              '�����ڲ�Connection
    
    frmDeliveryInvoice.Show vbModal
End Sub

Public Sub PurchasePlan(ByVal cnVal As ADODB.Connection)
'------------------------------
'����: ���òɹ��ƻ�������������
'------------------------------
    If cnVal Is Nothing Then
        MsgBox "����ZLHIS�����ݿ�����ʧ�ܣ���ֹ���룡", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    If gobjComLib Is Nothing Then
        MsgBox "����zl9ComLib����ʧ�ܣ���ֹ���룡", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    Set gcnOracle = cnVal

    Call gobjComLib.InitCommon(cnVal)              '�����ڲ�Connection
    
    frmPurchasePlan.Show vbModal
End Sub

Public Sub ClearDrugInfo(ByVal lngDrugID As Long, ByVal lngPurchaseID As Long)
'------------------------------
'����: ɾ��ҩƷ��Ϣ
'����:
'  lngDrugID: ҩƷID
'  lngPurchaseID: �б굥λID
'------------------------------
    Dim strDB As String, strServer As String, strUser As String, strPWD As String
    Dim strClear As String, strResultMESS As String
    Dim rsClear As New ADODB.Recordset
    
'    Set gcnOracle = cnVal
'    Call initCommon(cnVal)
    
    '���Ӳ���
    strDB = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="DBNAME", Default:="")
    strServer = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="SERVER", Default:="")
    strUser = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="USER", Default:="")
    strPWD = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="PASSWORD", Default:="")
    strPWD = StringEnDeCodecn(strPWD, 68)
    
    'Ĭ��MSSQL��ʽ����
    If MSSQLServerOpen(strServer, strDB, strUser, strPWD) = False Then
        If MsgBox("�Ƿ������������ӣ�", vbQuestion + vbYesNo + vbDefaultButton1, GSTR_MESSAGE) = vbNo Then
            Exit Sub
        Else
            frmOutsideLinkSet.Show vbModal
            If gblnSetupFinish = False Then Exit Sub
        End If
    End If
    
    strClear = "declare @msg varchar(2000) " & vbNewLine _
             & "execute p_ins_waredict '000006'" _
             & "  ,'" & IIf(lngPurchaseID = 0, "", lngPurchaseID) & "'" _
             & "  ,''" _
             & "  ,'" & IIf(lngDrugID = 0, "", lngDrugID) & "'" _
             & "  ,''" _
             & "  ,''" _
             & "  ,''" _
             & "  ,''" _
             & "  ,''" _
             & "  ,0" _
             & "  ,''" _
             & "  ,''" _
             & "  ,1 ,@msg output " & vbNewLine _
             & "select @msg result_msg "
    On Error Resume Next
    rsClear.Open strClear, gcnOutside
    strResultMESS = rsClear!result_msg
    rsClear.Close
    On Error GoTo 0
    
    If InStr(strResultMESS, "�ɹ�") = 0 Or strResultMESS = "" Then
        If strResultMESS = "" Then
            MsgBox "ͬ������ƽ̨ҩƷ��Ϣʧ�ܣ�", vbInformation, GSTR_MESSAGE
        Else
            MsgBox "ͬ������ƽ̨ҩƷ��Ϣ��" & vbNewLine & strResultMESS, vbInformation, GSTR_MESSAGE
        End If
    End If
    
End Sub

Public Sub UploadDrugInfo(ByVal frmOwner As Object, ByVal cnVal As ADODB.Connection, Optional ByVal lngDrugID As Long)
'------------------------------
'����: �ϴ�ҩƷ��Ϣ
'����:
'  frmOwner:  �����������
'  lngDrugID: ҩƷID    Ϊ0ʱ���ϴ�ȫ��ҩƷ��Ϣ
'------------------------------
    Dim strDB As String, strServer As String, strUser As String, strPWD As String
    Dim strSQL As String, strUpload As String, strResultMESS As String
    Dim rsTemp As ADODB.Recordset, rsUpload As New ADODB.Recordset
    Dim lngFailedQTY As Long
    
    If cnVal Is Nothing Then
        MsgBox "����ZLHIS�����ݿ�����ʧ�ܣ���ֹ���룡", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    If gobjComLib Is Nothing Then
        MsgBox "����zl9ComLib����ʧ�ܣ���ֹ���룡", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    Set gcnOracle = cnVal
    Call gobjComLib.InitCommon(cnVal)
    
    '���Ӳ���
    strDB = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="DBNAME", Default:="")
    strServer = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="SERVER", Default:="")
    strUser = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="USER", Default:="")
    strPWD = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="PASSWORD", Default:="")
    strPWD = StringEnDeCodecn(strPWD, 68)
    
    'Ĭ��MSSQL��ʽ����
    If MSSQLServerOpen(strServer, strDB, strUser, strPWD) = False Then
        If MsgBox("�Ƿ������������ӣ�", vbQuestion + vbYesNo + vbDefaultButton1, GSTR_MESSAGE) = vbNo Then
            Exit Sub
        Else
            frmOutsideLinkSet.Show vbModal
            If gblnSetupFinish = False Then Exit Sub
        End If
    End If
    
    strSQL = "Select a.Id, a.����, a.���, a.����, b.ҩ�ⵥλ, Round(b.�ɱ��� * b.ҩ���װ, 4) ҩ�ⵥ��, b.��׼�ĺ�, c.���� ͨ����, " _
           & "  c.ƴ����, c.��ѧ��, d.��λid ��Ӧ�̴���, e.���� ��Ӧ�� " _
           & "From �շ���ĿĿ¼ A, ҩƷ��� B," _
           & "  (Select �շ�ϸĿid, Min(Decode(����, 1, ����)) ����, Min(Decode(����, 1, ����)) ƴ����, Min(Decode(����, 0, ����)) ��ѧ�� " _
           & "   From �շ���Ŀ���� " _
           & "   Where ���� In (0, 1) And ���� In (1, 2) " _
           & "   Group By �շ�ϸĿid) C, " _
           & "  (Select ҩƷid, ��λid From ҩƷ�б굥λ Where ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) D, " _
           & "  (Select ID, ���� From ��Ӧ�� Where ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') And ĩ�� = 1 And Substr(����, 1, 1) = '1') E " _
           & "Where a.Id = b.ҩƷid And a.Id = c.�շ�ϸĿid(+) And b.ҩƷid = d.ҩƷid(+) And d.��λid = e.Id(+) " _
           & "  And a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') And a.��� In ('5', '6') "
    If lngDrugID > 0 Then
        strSQL = strSQL & " and a.ID=[1] "
    End If
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "�ϴ�ҩƷ��Ϣ������ƽ̨", lngDrugID)
    
    If Not rsTemp.EOF And lngDrugID = 0 Then
        With frmMessage
            .Show , frmOwner
            .Caption = "�����ϴ�ҩƷ��Ϣ"
            .lblMess.Caption = ""
            .cmdClose.Enabled = False
            .pbrMess.Min = 0: .pbrMess.Max = rsTemp.RecordCount: .pbrMess.Value = 0
        End With
        rsTemp.MoveFirst
        '���ҩƷ��Ϣ
        strUpload = "declare @msg varchar(2000) " & vbNewLine _
                  & "execute p_ins_waredict '000006'" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,0 " _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,-1 ,@msg output " & vbNewLine _
                  & "select @msg result_msg "
        On Error Resume Next
        rsUpload.Open strUpload, gcnOutside
        rsUpload.Close
        On Error GoTo 0
    End If
    
    Do While Not rsTemp.EOF
        strResultMESS = ""
        strUpload = "declare @msg varchar(2000) " & vbNewLine _
                  & "execute p_ins_waredict '000006'" _
                  & "  ,'" & rsTemp!��Ӧ�̴��� & "'" _
                  & "  ,'" & rsTemp!��Ӧ�� & "'" _
                  & "  ,'" & rsTemp!Id & "'" _
                  & "  ,'" & rsTemp!��ѧ�� & "'" _
                  & "  ,'" & rsTemp!ͨ���� & "'" _
                  & "  ,'" & rsTemp!ƴ���� & "'" _
                  & "  ,'" & rsTemp!��� & "'" _
                  & "  ,'" & rsTemp!ҩ�ⵥλ & "'" _
                  & "  ," & rsTemp!ҩ�ⵥ�� _
                  & "  ,'" & rsTemp!���� & "'" _
                  & "  ,'" & rsTemp!��׼�ĺ� & "'" _
                  & "  ,0 ,@msg output " & vbNewLine _
                  & "select @msg result_msg "
        On Error Resume Next
        rsUpload.Open strUpload, gcnOutside
        DoEvents
        strResultMESS = rsUpload!result_msg
        rsUpload.Close
        On Error GoTo 0
        '������ʾ��Ϣ
        If InStr(strResultMESS, "�ɹ�") = 0 Or strResultMESS = "" Then
            If lngDrugID = 0 Then
                '�����ϴ�
                lngFailedQTY = lngFailedQTY + 1
                With frmMessage
                    .txtContents.Text = "�ϴ�ҩƷ���룺" & rsTemp!���� & "�� ͨ������" & rsTemp!ͨ���� & " ��" & vbNewLine _
                                      & "    " & IIf(strResultMESS = "", "�ϴ�ʧ�ܣ�", strResultMESS) & vbNewLine _
                                      & .txtContents.Text
                End With
            Else
                '�����ϴ�
                MsgBox "�ϴ�ʧ�ܣ�", vbInformation, GSTR_MESSAGE
            End If
            Err.Clear
        End If
        If lngDrugID = 0 Then
            With frmMessage
                .pbrMess.Value = rsTemp.AbsolutePosition
                .lblMess = "�ܹ� " & .pbrMess.Max & " ����¼��δ�ϴ� " & lngFailedQTY + .pbrMess.Max - .pbrMess.Value & " ����¼��"
            End With
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If lngDrugID = 0 Then
        frmMessage.txtContents.Text = "�����ϴ���ɣ�" & vbNewLine & vbNewLine & frmMessage.txtContents.Text
        frmMessage.cmdClose.Enabled = True
    End If
End Sub

Private Sub Class_Initialize()
    '����zl9ComLib����
    On Error Resume Next
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
End Sub

Private Sub Class_Terminate()
    Set gobjComLib = Nothing
End Sub