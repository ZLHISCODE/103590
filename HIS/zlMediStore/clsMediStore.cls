VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMediStore"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit
  Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, FrmMain As Object, strDbUser As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngSys : ϵͳ���
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '���أ�
    '------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strDept As String
    Dim strPrivs As String
    
    Set gcnOracle = cnMain
    
    glngSys = lngSys
    
'    Set frmWin = FrmMain
    gstrAviPath = GetSetting("ZLSOFT", "ע����Ϣ", "gstrAviPath", "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrVersion = GetSetting("ZLSOFT", "ע����Ϣ", "gstrVersion", "")
    
    gstrMatchMethod = GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0)
    gbytSimpleCodeTrans = Val(zldatabase.GetPara("����ƥ�䷽ʽ�л�", , , 1))
    
    glngModul = lngModul
    gstrprivs = GetPrivFunc(lngSys, lngModul)
'    zlbillprint.BillInit gcnOracle
    
     '��ȡ��¼�û���Ϣ
    gstrDbUser = strDbUser   '  GetSetting("ZLHIS", "��¼��Ϣ", frmMain.hwnd, GetSetting("ZLHIS", "��¼��Ϣ", "USER", ""))
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not CheckValid Then Exit Sub
    
    '��ȡҩƷ��ͨ����Ҫʹ�õ�ϵͳ����ֵ
    Call GetSysParms
    
    'ȡҩƷ�۸�ȼ�
    Call GetPriceClass
    
    '-------------------------------------------------
    '1300  ҩƷ�⹺������
    '1301  ҩƷ����������
    '1302  ҩƷ����������
    '1303  ҩƷ����۵�������
    '1304  ҩƷ�ƿ����
    '1305  ҩƷ���ù���
    '1306  ҩƷ�����������
    '1307  ҩƷ�̵����
    '1308  ҩƷ�������
    '1309  ҩƷ����ѯ
    '1320  ҩƷ�������
    
    '1330  ҩƷ�ƻ�����
    '1331  ҩƷ��������
    '1332  ҩƷ������
    '1333  ҩƷ���۹���
    '2287  ҩƷ�������
    
    Dim frmPurchaseList As Form
    Dim frmTemp As Form
    Dim strTitle As String
    
    
    On Error Resume Next
    
    Call setNOtExcetePrice '��ִ�е�ʱ�仹û��ִ�еļ۸�
    Select Case lngModul
        Case 1300
            strTitle = "ҩƷ�⹺������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
            
             '.Show , frmMain            'ҩƷ�⹺������
        Case 1301
            strTitle = "ҩƷ����������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        Case 1302                        'ҩƷ����������
            strTitle = "ҩƷ����������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
'        Case 1303
'            strTitle = "����۵�������"
'            For Each frmTemp In Forms
'                If frmTemp.Caption = strTitle Then
'                    frmTemp.Show
'                    frmTemp.ZOrder 0
'                    Exit Sub
'                End If
'            Next
'            Set frmPurchaseList = New frmMainList
'            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        Case 1304
            strTitle = "ҩƷ�ƿ����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        
        Case 1305
            strTitle = "ҩƷ���ù���"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        Case 1306
            strTitle = "ҩƷ�����������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        Case 1307
            strTitle = "ҩƷ�̵����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmNewCheckMain
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        Case 1320
            'frmDrugPaymentList.Show , FrmMain
        Case 1321
            'frmӦ�����ѯ.Show , FrmMain
        Case 1330
            frmDrugPlanList.Show , FrmMain
        Case 1331
            frmDrugQualityList.Show , FrmMain
'        Case 1308
'            frmDiffPriceRecalCard.Show , FrmMain
        Case 1309
            frmDrugQuery.Show , FrmMain
        Case 1332
            frmClosingAccount.Show , FrmMain
        Case 1333
            frmMediPrice.Show , FrmMain
        Case 1348 'ҩƷ�������
            frmMediCheckList.Show , FrmMain
    End Select
End Sub

Private Sub setNOtExcetePrice()
    '����ʱ�仹δִ�е���ҩƷִ�е���
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct i.Id As ҩƷid " & _
               " From �շ���ĿĿ¼ I, �շѼ�Ŀ N, ҩƷ��� P" & _
               " Where i.Id = n.�շ�ϸĿid And i.Id = p.ҩƷid And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & _
                   " n.�䶯ԭ�� = 0 And Sysdate>n.ִ������" & GetPriceClassString("N") & _
               " Union " & _
               " Select Distinct a.ҩƷid From ҩƷ�۸��¼ A Where a.��¼״̬ = 0 And a.ִ������ <= Sysdate " & _
               " Order By ҩƷid "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ִ�е���")
    
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    For i = 0 To rsTemp.RecordCount - 1
        gstrSQL = "Zl_ҩƷ�շ���¼_Adjust(" & rsTemp!ҩƷID & ")"
        zldatabase.ExecuteProcedure gstrSQL, "ִ�е���"
        rsTemp.MoveNext
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, strDbUser As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngSys : ϵͳ���
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   lngMain:��������
    '���أ�
    '------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strDept As String
    Dim strPrivs As String
    
    Set gcnOracle = cnMain
    
    glngSys = lngSys
    
    gstrAviPath = GetSetting("ZLSOFT", "ע����Ϣ", "gstrAviPath", "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrVersion = GetSetting("ZLSOFT", "ע����Ϣ", "gstrVersion", "")
    
    gstrMatchMethod = GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0)
    gbytSimpleCodeTrans = Val(zldatabase.GetPara("����ƥ�䷽ʽ�л�", , , 1))
    
    glngModul = lngModul
    gstrprivs = GetPrivFunc(lngSys, lngModul)
    
    '��ȡ��¼�û���Ϣ
    gstrDbUser = strDbUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not CheckValid Then Exit Sub
    
    '��ȡҩƷ��ͨ����Ҫʹ�õ�ϵͳ����ֵ
    Call GetSysParms
   
    '-------------------------------------------------
    '1300  ҩƷ�⹺������
    '1301  ҩƷ����������
    '1302  ҩƷ����������
    '1303  ҩƷ����۵�������
    '1304  ҩƷ�ƿ����
    '1305  ҩƷ���ù���
    '1306  ҩƷ�����������
    '1307  ҩƷ�̵����
    '1308  ҩƷ�������
    '1309  ҩƷ����ѯ
    '1320  ҩƷ�������
    
    '1330  ҩƷ�ƻ�����
    '1331  ҩƷ��������
    '1332  ҩƷ������
    '1333  ҩƷ���۹���
    
    Dim frmPurchaseList As Form
    Dim frmTemp As Form
    Dim strTitle As String
    
    On Error Resume Next
    
    Call setNOtExcetePrice '��ִ�е�ʱ�仹û��ִ�еļ۸�
    Select Case lngModul
        Case 1300 'ҩƷ�⹺������
            strTitle = "ҩƷ�⹺������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1301
            strTitle = "ҩƷ����������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1302                        'ҩƷ����������
            strTitle = "ҩƷ����������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
'        Case 1303
'            strTitle = "����۵�������"
'            For Each frmTemp In Forms
'                If frmTemp.Caption = strTitle Then
'                    frmTemp.Show
'                    frmTemp.ZOrder 0
'                    Exit Sub
'                End If
'            Next
'            Set frmPurchaseList = New frmMainList
'            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1304
            strTitle = "ҩƷ�ƿ����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        
        Case 1305
            strTitle = "ҩƷ���ù���"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1306
            strTitle = "ҩƷ�����������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1307
            strTitle = "ҩƷ�̵����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmNewCheckMain
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1320
            'OS.ShowChildWindow frmDrugPaymentList.hWnd, lngMain
        Case 1321
            'OS.ShowChildWindow frmӦ�����ѯ.hWnd, lngMain
        Case 1330
            OS.ShowChildWindow frmDrugPlanList.hWnd, lngMain
            
        Case 1331
            OS.ShowChildWindow frmDrugQualityList.hWnd, lngMain
'        Case 1308
'            OS.ShowChildWindow frmDiffPriceRecalCard.hWnd, lngMain
        Case 1309
            OS.ShowChildWindow frmDrugQuery.hWnd, lngMain
        Case 1332
            OS.ShowChildWindow frmClosingAccount.hWnd, lngMain
        Case 1333
            OS.ShowChildWindow frmMediPrice.hWnd, lngMain
        Case 1348 'ҩƷ�������
            OS.ShowChildWindow frmMediCheckList.hWnd, lngMain
    End Select
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.count = 0)
End Function
