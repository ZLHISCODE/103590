VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Begin VB.Form frmAddPatient 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "���˱걾����"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraִ��ѡ�� 
      Caption         =   "ִ��ѡ��"
      Height          =   975
      Left            =   90
      TabIndex        =   25
      Top             =   6120
      Width           =   11565
      Begin MSComCtl2.DTPicker DTP 
         Height          =   285
         Left            =   1020
         TabIndex        =   49
         Top             =   570
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   245628931
         CurrentDate     =   39620
      End
      Begin VB.TextBox txt���ɱ걾�� 
         Height          =   270
         Left            =   6780
         TabIndex        =   47
         Top             =   570
         Width           =   1725
      End
      Begin VB.TextBox txt�걾���� 
         Height          =   270
         Left            =   4050
         TabIndex        =   46
         Top             =   570
         Width           =   1725
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         Left            =   9630
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   210
         Width           =   1725
      End
      Begin VB.ComboBox cboִ�п��� 
         Height          =   300
         Left            =   6780
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   210
         Width           =   1725
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   210
         Width           =   2145
      End
      Begin VB.ComboBox cbo����ҽ�� 
         Height          =   300
         Left            =   4050
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   210
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   210
         TabIndex        =   48
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl�걾���� 
         Caption         =   "�걾����"
         Height          =   195
         Left            =   3240
         TabIndex        =   37
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lbl���ɱ걾�� 
         Caption         =   "�� �� ��"
         Height          =   195
         Left            =   5940
         TabIndex        =   35
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   8790
         TabIndex        =   33
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblִ�п��� 
         AutoSize        =   -1  'True
         Caption         =   "ִ�п���"
         Height          =   180
         Left            =   5925
         TabIndex        =   31
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   210
         TabIndex        =   29
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl����ҽ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ҽ��"
         Height          =   180
         Left            =   3255
         TabIndex        =   27
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10110
      TabIndex        =   14
      Top             =   7260
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8580
      TabIndex        =   13
      Top             =   7260
      Width           =   1100
   End
   Begin VB.Frame fra��Ŀѡ�� 
      Caption         =   "��Ŀѡ��"
      Height          =   2805
      Left            =   90
      TabIndex        =   6
      Top             =   3270
      Width           =   11565
      Begin XtremeReportControl.ReportControl rptItemSelect 
         Height          =   2445
         Left            =   7170
         TabIndex        =   12
         Top             =   210
         Width           =   4305
         _Version        =   589884
         _ExtentX        =   7594
         _ExtentY        =   4313
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin XtremeReportControl.ReportControl rptItemSource 
         Height          =   2445
         Left            =   2460
         TabIndex        =   7
         Top             =   210
         Width           =   3975
         _Version        =   589884
         _ExtentX        =   7011
         _ExtentY        =   4313
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "&P"
         Height          =   255
         Left            =   2010
         TabIndex        =   32
         Top             =   765
         Width           =   285
      End
      Begin VB.CommandButton cmdItemLeftAll 
         Caption         =   "<<"
         Height          =   375
         Left            =   6630
         TabIndex        =   24
         Top             =   2130
         Width           =   405
      End
      Begin VB.CommandButton cmdItemLeft 
         Caption         =   "<"
         Height          =   375
         Left            =   6630
         TabIndex        =   23
         Top             =   1590
         Width           =   405
      End
      Begin VB.CommandButton cmdItemRightAll 
         Caption         =   ">>"
         Height          =   375
         Left            =   6630
         TabIndex        =   22
         Top             =   1050
         Width           =   405
      End
      Begin VB.CommandButton cmdItemRight 
         Caption         =   ">"
         Height          =   375
         Left            =   6630
         TabIndex        =   21
         Top             =   540
         Width           =   405
      End
      Begin VB.OptionButton optIF 
         Caption         =   "����"
         Height          =   225
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Top             =   1200
         Width           =   765
      End
      Begin VB.OptionButton optIF 
         Caption         =   "���"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   1200
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.TextBox txt���� 
         Height          =   270
         Left            =   750
         TabIndex        =   11
         Top             =   780
         Width           =   1275
      End
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbl���� 
         Caption         =   "��  ��"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   825
         Width           =   615
      End
      Begin VB.Label lbl��� 
         Caption         =   "��  ��"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   420
         Width           =   555
      End
   End
   Begin VB.Frame fra����ѡ�� 
      Caption         =   "����ѡ��"
      Height          =   3165
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   11565
      Begin XtremeReportControl.ReportControl rptListSelect 
         Height          =   2865
         Left            =   7170
         TabIndex        =   2
         Top             =   180
         Width           =   4275
         _Version        =   589884
         _ExtentX        =   7541
         _ExtentY        =   5054
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin XtremeReportControl.ReportControl rptListSource 
         Height          =   2865
         Left            =   2460
         TabIndex        =   1
         Top             =   180
         Width           =   3975
         _Version        =   589884
         _ExtentX        =   7011
         _ExtentY        =   5054
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   660
         Width           =   1575
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   1230
         TabIndex        =   44
         Top             =   2640
         Width           =   1100
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   270
         Width           =   1575
      End
      Begin VB.TextBox txt���� 
         Height          =   270
         Left            =   780
         TabIndex        =   40
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox txt�걾�� 
         Height          =   270
         Left            =   780
         TabIndex        =   38
         Top             =   1500
         Width           =   1545
      End
      Begin VB.CommandButton cmdPatientLeftAll 
         Caption         =   "<<"
         Height          =   375
         Left            =   6600
         TabIndex        =   20
         Top             =   2100
         Width           =   405
      End
      Begin VB.CommandButton cmdPatientLeft 
         Caption         =   "<"
         Height          =   375
         Left            =   6600
         TabIndex        =   19
         Top             =   1560
         Width           =   405
      End
      Begin VB.CommandButton cmdPatientRightAll 
         Caption         =   ">>"
         Height          =   375
         Left            =   6600
         TabIndex        =   18
         Top             =   1020
         Width           =   405
      End
      Begin VB.CommandButton cmdPatientRight 
         Caption         =   ">"
         Height          =   375
         Left            =   6600
         TabIndex        =   17
         Top             =   510
         Width           =   405
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   255
         Left            =   780
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   245628929
         CurrentDate     =   39533
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   255
         Left            =   780
         TabIndex        =   5
         Top             =   2250
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   245628929
         CurrentDate     =   39533
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "��  ��"
         Height          =   180
         Left            =   150
         TabIndex        =   36
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lbl���� 
         Caption         =   "��  ��"
         Height          =   195
         Left            =   150
         TabIndex        =   43
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lbl���� 
         Caption         =   "��  ��"
         Height          =   195
         Left            =   150
         TabIndex        =   41
         Top             =   1125
         Width           =   825
      End
      Begin VB.Label lbl�걾�� 
         Caption         =   "�걾��"
         Height          =   195
         Left            =   150
         TabIndex        =   39
         Top             =   1545
         Width           =   555
      End
      Begin VB.Label lbl���� 
         Caption         =   "��  ��"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   1950
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmAddPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlngMachineID As Long           '����ID
Dim mlngExecDeptID As Long          'ִ�п���ID

Private Enum mColP
    ����ID = 0: �걾���: ����: �Ա�: ����: Ӥ��: �Һŵ�: �����: סԺ��: ��������: ��ҳID: ��ʶ��: ����: ���˿���: ������: �������ID: ������Դ
End Enum
Private Enum mColI
    ID = 0: ����: ����: �걾: ��Ŀ���
End Enum

Private Sub Option2_Click()

End Sub

Private Sub cbo��������_Click()
    If Me.cbo��������.ItemData(Me.cbo��������.ListIndex) = -1 Then
        txt�걾����.Enabled = True
    Else
        txt�걾����.Text = ""
        txt�걾����.Enabled = False
    End If
End Sub

Private Sub cbo��������_Click()
    Dim lngApplyDept As Long
    Dim rsTmp As New ADODB.Recordset
    Dim lngKey As Long
    
    lngApplyDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
    lngKey = zldatabase.GetPara("frmAddPatient_����ҽ��", 100, 1208, -1)
    
    '�����Ӧ�����µ���Ա
    gstrSql = "select distinct a.id,a.���,a.���� from ��Ա�� a , ��Ա����˵�� b , ������Ա c " & _
                 " where a.id = b.��Աid and a.id = c.��ԱID and  b.��Ա���� in ('ҽ��','��ʿ') " & _
                 " and (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "
    If lngApplyDept = 0 Then
        gstrSql = gstrSql & " order by a.���"
    Else
        gstrSql = gstrSql & " and c.����id = [1] order by a.��� "
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, gstrSysName, lngApplyDept)
    
    With Me.cbo����ҽ��
        .Clear
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("���")) & "-" & Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
End Sub

Private Sub cbo����_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim lngDept As Long
    Dim lngKey As Long
    
    lngKey = zldatabase.GetPara("frmAddPatient_ѡ������", 100, 1208, -1)
    gstrSql = "select ID,����,����  from �������� "
    
    lngDept = cbo����.ItemData(cbo����.ListIndex)
    If lngDept > 0 Then
        gstrSql = gstrSql & " Where ʹ��С��ID = [1] "
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDept)
    With cbo����
        .Clear
        .AddItem "�ֹ�"
        .ItemData(.NewIndex) = -1
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
End Sub

Private Sub cboִ�п���_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim lngDept As Long
    Dim lngKey As Long
    
    lngKey = zldatabase.GetPara("frmAddPatient_��������", 100, 1208, -1)
    gstrSql = "select ID,����,����  from �������� "
    If cboִ�п���.ListIndex >= 0 Then
        lngDept = cboִ�п���.ItemData(cboִ�п���.ListIndex)
    End If
    If lngDept > 0 Then
        gstrSql = gstrSql & " Where ʹ��С��ID = [1] "
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDept)
    With cbo��������
        .Clear
        .AddItem "�ֹ�"
        .ItemData(.NewIndex) = -1
        If lngKey = -1 Then .ListIndex = .NewIndex
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdItemLeft_Click()
    MoveItem 2, 1, False
End Sub

Private Sub cmdItemLeftAll_Click()
    MoveItem 2, 1, True
End Sub

Private Sub cmdItemRight_Click()
    MoveItem 2, 2, False
End Sub

Private Sub cmdItemRightAll_Click()
    MoveItem 2, 2, True
End Sub

Private Sub cmdOk_Click()
    If chkSaveData = True Then
        Call SaveData
    End If
End Sub

Private Sub cmdPatientLeft_Click()
    MoveItem 1, 1, False
End Sub

Private Sub cmdPatientLeftAll_Click()
    MoveItem 1, 1, True
End Sub

Private Sub cmdPatientRight_Click()
    MoveItem 1, 2, False
End Sub

Private Sub cmdPatientRightAll_Click()
    MoveItem 1, 2, True
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strType As String
    Dim Record As ReportRecord
    Dim intLoop As Integer
    
    gstrSql = "Select Distinct A.ID, A.����, B.����, �걾��λ, D.��Ŀ���" & vbNewLine & _
                "From ������ĿĿ¼ A, ������Ŀ���� B, ���鱨����Ŀ C, ������Ŀ D" & vbNewLine & _
                "Where A.ID = B.������Ŀid And A.��� = 'C' And A.����Ӧ�� = 1 And A.ID = C.������Ŀid And " & vbNewLine & _
                " C.������Ŀid = D.������Ŀid(+) And D.��Ŀ��� In (1, 2,4) " & vbNewLine & _
                " And (A.����ʱ�� Is Null Or To_Char(A.����ʱ��, 'yyyy-mm-dd') = '3000-01-01') "
    
    If Me.cbo���.ListIndex > 0 Then
        gstrSql = gstrSql & " And a.�������� = [2] "
        strType = Mid(Me.cbo���.Text, InStr(Me.cbo���, "-") + 1)
    End If
    
    If Trim(Me.txt����.Text) <> "" Then
        gstrSql = gstrSql & " And (b.���� like [3] or d.��д like [3]) "
    End If
    
    gstrSql = gstrSql & " And nvl(a.�����Ŀ,0) = [4] "
   
    gstrSql = gstrSql & " Order by a.���� "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.optIF(0).Value, 1, 0), strType, "%" & UCase(Me.txt����) & "%", _
                            IIf(optIF(0).Value, 1, 0))
    rptItemSource.Records.DeleteAll:    rptItemSelect.Records.DeleteAll
    Do Until rsTmp.EOF
        Set Record = Me.rptItemSource.Records.Add
        For intLoop = 0 To Me.rptItemSource.Columns.Count
            Record.AddItem ""
        Next
        Record.Item(mColI.ID).Value = Nvl(rsTmp("ID"))
        Record.Item(mColI.����).Value = Nvl(rsTmp("����"))
        Record.Item(mColI.����).Value = Nvl(rsTmp("����"))
        Record.Item(mColI.�걾).Value = Nvl(rsTmp("�걾��λ"))
        Record.Item(mColI.��Ŀ���).Value = Nvl(rsTmp("��Ŀ���"))
        rsTmp.MoveNext
    Loop
    rptItemSource.Populate: rptItemSelect.Populate
    Me.txt����.Text = ""
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strWhere As String
    Dim lngDept As Long, lngMachine As Long
    Dim strBegingNO As String, strEndNO As String
    Dim astrItem() As String
    Dim Record As ReportRecord                                      '�б��¼��
    
    On Error GoTo errH
    
    gstrSql = "Select to_Number(�걾���) as �걾���,����id, Ӥ��,����, �Ա�, ����, �Һŵ�, �����, סԺ��, ��������, ��ҳid, ��ʶ��, " & vbNewLine & _
              " ����, ���˿���, ������, �������id,������Դ, " & vbNewLine & _
              " Decode(����id, Null," & vbNewLine & _
              "                 To_Char(Trunc(�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(�걾���, 10000), '0000')," & vbNewLine & _
              "                 �걾���) As �걾����ʾ " & vbNewLine & _
              " From ����걾��¼ where ҽ��ID is not null and  ����ʱ�� between [1] and [2] "
    'ִ�п���
    If cbo����.Text <> "" Then
        lngDept = cbo����.ItemData(cbo����.ListIndex)
        If lngDept > 0 Then
            strWhere = " And ִ�п���ID = [3] "
        End If
    End If
    '����
    If cbo����.Text <> "" Then
        lngMachine = cbo����.ItemData(cbo����.ListIndex)
        If lngMachine = -1 Then
            strWhere = strWhere & " And ����ID is null "
        ElseIf lngMachine > 0 Then
            strWhere = strWhere & " And ����ID = [4]"
        End If
    End If
    '�걾��
    If Trim(txt�걾��) <> "" Then
        txt�걾�� = Replace(Replace(txt�걾��, "��", "~"), "-", "~")
        varItem = Split(Trim(txt�걾��.Text), ",")
        
        For lngLoop = 0 To UBound(varItem)
            astrItem = Split(varItem(lngLoop), "~")
            
            If UBound(astrItem) <= 0 Then
                strBegingNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
            Else
                strBegingNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(astrItem(0)), Val(astrItem(0))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(astrItem(1)), Val(astrItem(1))))
            End If
            If lngLoop = 0 Then
                strWhere = strWhere & " and (to_Number(�걾���) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            Else
                strWhere = strWhere & "  or to_Number(�걾���) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            End If
        Next
        If lngLoop >= 0 Then strWhere = strWhere & ")"
    ElseIf Trim(txt����) <> "" Then
        strWhere = strWhere & " and to_Number(�걾���) between [5] and [6] "
        strBegingNO = TransSampleNO(Val(Me.txt����) & "-0001")
        strEndNO = TransSampleNO(Val(Me.txt����) & "-9999")
    End If
    
    gstrSql = gstrSql & strWhere & " Order by to_Number(�걾���) "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(Format(dtpBegin.Value, "yyyy-mm-dd 00:00:00")), _
                CDate(Format(dtpEnd.Value, "yyyy-mm-dd 23:59:59")), lngDept, lngMachine, Val(strBegingNO), Val(strEndNO))
    Me.rptListSelect.Records.DeleteAll
    Me.rptListSource.Records.DeleteAll
    Me.rptListSource.Populate
    Me.rptListSelect.Populate
    Do Until rsTmp.EOF
        
        Set Record = Me.rptListSource.Records.Add
            For intLoop = 0 To Me.rptListSource.Columns.Count
                Record.AddItem ""
            Next
            Record.Item(mColP.����ID).Value = Nvl(rsTmp("����ID"))
            Record.Item(mColP.�걾���).Value = Nvl(rsTmp("�걾���"))
            Record.Item(mColP.�걾���).Caption = Nvl(rsTmp("�걾����ʾ"))
            Record.Item(mColP.����).Value = Nvl(rsTmp("����"))
            Record.Item(mColP.�Ա�).Value = Nvl(rsTmp("�Ա�"))
            Record.Item(mColP.����).Value = Nvl(rsTmp("����"))
            Record.Item(mColP.Ӥ��).Value = Nvl(rsTmp("Ӥ��"))
            Record.Item(mColP.�Һŵ�).Value = Nvl(rsTmp("�Һŵ�"))
            Record.Item(mColP.�����).Value = Nvl(rsTmp("�����"))
            Record.Item(mColP.סԺ��).Value = Nvl(rsTmp("סԺ��"))
            Record.Item(mColP.��������).Value = Nvl(rsTmp("��������"))
            Record.Item(mColP.��ҳID).Value = Nvl(rsTmp("��ҳID"))
            Record.Item(mColP.��ʶ��).Value = Nvl(rsTmp("��ʶ��"))
            Record.Item(mColP.����).Value = Nvl(rsTmp("����"))
            Record.Item(mColP.���˿���).Value = Nvl(rsTmp("���˿���"))
            Record.Item(mColP.������).Value = Nvl(rsTmp("������"))
            Record.Item(mColP.�������ID).Value = Nvl(rsTmp("�������ID"))
            Record.Item(mColP.������Դ).Value = Nvl(rsTmp("������Դ"))
        rsTmp.MoveNext
    Loop
    Me.rptListSource.Populate
    Me.rptListSelect.Populate
    
    Exit Sub
errH:
    If errcenter() = 1 Then
        Resume
    End If
    Call saveerrlog
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strDate As String
    Dim Column As ReportColumn
    Dim lngKey As Long
    
    DTP.Value = Now
    
    '��������������
    '==���� �� ִ�п���
    lngKey = zldatabase.GetPara("frmAddPatient_ѡ�����", 100, 1208, -1)
    gstrSql = "select id,����,���� from ���ű� a , ��������˵�� b" & vbNewLine & _
              "where a.id = b.����id and �������� = '����'" & vbNewLine & _
              "order by ����"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    With cbo����
        .AddItem "���п���"
        .ItemData(.NewIndex) = 0
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
    
    lngKey = zldatabase.GetPara("frmAddPatient_ִ�п���", 100, 1208, -1)
    rsTmp.MoveFirst
    With cboִ�п���
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
        
    '==����
    strDate = zldatabase.Currentdate
    Me.dtpBegin = strDate
    Me.dtpEnd = strDate
    
    '==���
    lngKey = zldatabase.GetPara("frmAddPatient_ѡ�����", 100, 1208, -1)
    gstrSql = "select ����,���� from ���Ƽ������� "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With cbo���
        .AddItem "�������"
        .ItemData(.NewIndex) = 0
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = rsTmp("����")
            If lngKey = rsTmp("����") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
    
    '==��������
    lngKey = zldatabase.GetPara("frmAddPatient_��������", 100, 1208, -1)
    gstrSql = "select distinct a.id,a.����,a.���� from ���ű� a , ��������˵�� b " & _
                 " where a.id = b.����id and b.�������� in ('����','����','�ٴ�')  order by a.����"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With cbo��������
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
                
    '==��������
    gstrSql = "select ID,����,����  from �������� "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With cbo��������
        .AddItem "�ֹ�"
        .ItemData(.NewIndex) = -1
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = rsTmp("ID")
            rsTmp.MoveNext
        Loop
    End With
    
    '==�������
    gstrSql = "select ����,���� from ���Ƽ������� order by ����"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With cbo���
        .Clear
        .AddItem "�������"
        .ItemData(.NewIndex) = 0
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = Nvl(rsTmp("����"))
            rsTmp.MoveNext
        Loop
    End With
    
    With Me.rptListSource.Columns
        rptListSource.AllowColumnRemove = False
        rptListSource.ShowItemsInGroups = False
        
        With rptListSource.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mColP.����ID, "����ID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.�걾���, "�걾���", 50, True)
        Set Column = .Add(mColP.����, "����", 75, True)
        Set Column = .Add(mColP.�Ա�, "�Ա�", 45, True)
        Set Column = .Add(mColP.����, "����", 60, True)
        Set Column = .Add(mColP.Ӥ��, "Ӥ��", 45, True)
        Set Column = .Add(mColP.�Һŵ�, "�Һŵ�", 75, True): Column.Visible = False
        Set Column = .Add(mColP.�����, "�����", 75, True): Column.Visible = False
        Set Column = .Add(mColP.סԺ��, "סԺ��", 75, True): Column.Visible = False
        Set Column = .Add(mColP.��������, "��������", 75, True): Column.Visible = False
        Set Column = .Add(mColP.��ҳID, "��ҳID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.��ʶ��, "��ʶ��", 75, True): Column.Visible = False
        Set Column = .Add(mColP.����, "����", 75, True): Column.Visible = False
        Set Column = .Add(mColP.���˿���, "���˿���", 75, True): Column.Visible = False
        Set Column = .Add(mColP.������, "������", 75, True): Column.Visible = False
        Set Column = .Add(mColP.�������ID, "�������ID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.������Դ, "������Դ", 75, True): Column.Visible = False
    End With
    
    With Me.rptListSelect.Columns
        rptListSelect.AllowColumnRemove = False
        rptListSelect.ShowItemsInGroups = False
        
        With rptListSelect.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mColP.����ID, "����ID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.�걾���, "�걾���", 50, True)
        Set Column = .Add(mColP.����, "����", 75, True)
        Set Column = .Add(mColP.�Ա�, "�Ա�", 45, True)
        Set Column = .Add(mColP.����, "����", 60, True)
        Set Column = .Add(mColP.Ӥ��, "Ӥ��", 45, True)
        Set Column = .Add(mColP.�Һŵ�, "�Һŵ�", 75, True): Column.Visible = False
        Set Column = .Add(mColP.�����, "�����", 75, True): Column.Visible = False
        Set Column = .Add(mColP.סԺ��, "סԺ��", 75, True): Column.Visible = False
        Set Column = .Add(mColP.��������, "��������", 75, True): Column.Visible = False
        Set Column = .Add(mColP.��ҳID, "��ҳID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.��ʶ��, "��ʶ��", 75, True): Column.Visible = False
        Set Column = .Add(mColP.����, "����", 75, True): Column.Visible = False
        Set Column = .Add(mColP.���˿���, "���˿���", 75, True): Column.Visible = False
        Set Column = .Add(mColP.������, "������", 75, True): Column.Visible = False
        Set Column = .Add(mColP.�������ID, "�������ID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.������Դ, "������Դ", 75, True): Column.Visible = False
    End With
    
    With Me.rptItemSource.Columns
        rptItemSource.AllowColumnRemove = False
        rptItemSource.ShowItemsInGroups = False
        
        With rptItemSource.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mColI.ID, "ID", 75, True): Column.Visible = False
        Set Column = .Add(mColI.����, "����", 75, True)
        Set Column = .Add(mColI.����, "����", 100, True)
        Set Column = .Add(mColI.�걾, "�걾", 60, True)
        Set Column = .Add(mColI.��Ŀ���, "��Ŀ���", 60, True): Column.Visible = False
    End With
    
    With Me.rptItemSelect.Columns
        rptItemSelect.AllowColumnRemove = False
        rptItemSelect.ShowItemsInGroups = False
        
        With rptItemSelect.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mColI.ID, "ID", 75, True): Column.Visible = False
        Set Column = .Add(mColI.����, "����", 75, True)
        Set Column = .Add(mColI.����, "����", 75, True)
        Set Column = .Add(mColI.�걾, "�걾", 60, True)
        Set Column = .Add(mColI.��Ŀ���, "��Ŀ���", 60, True): Column.Visible = False
    End With
End Sub
Public Sub ShowMe(objfrm As Object, lngMachineID As Long, lngExecDeptID As Long)
    mlngMachineID = lngMachineID
    mlngExecDeptID = lngExecDeptID
    Me.Show vbModal, objfrm
End Sub

Private Function chkRepeat(lngID As Long, rptRows As ReportRows, strItemIndex As Integer) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����           �����ظ�
    '����
    '����           True=�ظ�  False=���ظ�
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    For intLoop = 0 To rptRows.Count - 1
        If lngID = rptRows(intLoop).Record(strItemIndex).Value Then
            chkRepeat = True
            Exit Function
        End If
    Next
    chkrepeate = False
End Function
Private Sub MoveItem(intType As Integer, intMove As Integer, AllItem As Boolean)
    '����               �ƶ���Ŀ��ָ���б����
    '����               intType 1=���� 2=��Ŀ
    '                   intMove 1= Left 2=Right
    '                   AllItem True=���� False=����
    Dim Record As ReportRecord
    Dim intLoop As Integer, intRow As Integer
    Dim lngKey As Long
    
    '����
    If intType = 1 Then
        If intMove = 1 Then
            'Left
            If AllItem = False Then
                If Not Me.rptListSelect.FocusedRow Is Nothing Then
                    lngKey = Me.rptListSelect.FocusedRow.Record(mColP.����ID).Value
                    If chkRepeat(lngKey, Me.rptListSource.Rows, CInt(mColP.����ID)) = False Then
                        Set Record = Me.rptListSource.Records.Add
                        For intLoop = 0 To Me.rptListSelect.Columns.Count
                            Record.AddItem ""
                            Record.Item(intLoop).Value = Me.rptListSelect.FocusedRow.Record(intLoop).Value
                        Next
                        Me.rptListSelect.Records.RemoveAt (Me.rptListSelect.FocusedRow.Record.Index)
                    End If
                End If
            Else
                For intLoop = 0 To Me.rptListSelect.Records.Count - 1
                    lngKey = Me.rptListSelect.Records(intLoop).Item(mColP.����ID).Value
                    If chkRepeat(lngKey, Me.rptListSource.Rows, CInt(mColP.����ID)) = False Then
                        Set Record = Me.rptListSource.Records.Add
                        For intRow = 0 To Me.rptListSelect.Columns.Count
                            Record.AddItem ""
                            Record.Item(intRow).Value = Me.rptListSelect.Records(intLoop).Item(intRow).Value
                         Next
                    End If
                Next
                Me.rptListSelect.Records.DeleteAll
            End If
            Me.rptListSelect.Populate: Me.rptListSource.Populate
        Else
            'Right
            If AllItem = False Then
                If Not Me.rptListSource.FocusedRow Is Nothing Then
                    lngKey = Me.rptListSource.FocusedRow.Record(mColP.����ID).Value
                    If chkRepeat(lngKey, Me.rptListSelect.Rows, CInt(mColP.����ID)) = False Then
                        Set Record = Me.rptListSelect.Records.Add
                        For intLoop = 0 To Me.rptListSource.Columns.Count
                            Record.AddItem ""
                            Record.Item(intLoop).Value = Me.rptListSource.FocusedRow.Record(intLoop).Value
                        Next
                        Me.rptListSource.Records.RemoveAt (Me.rptListSource.FocusedRow.Record.Index)
                        
                    End If
                End If
            Else
                For intLoop = 0 To Me.rptListSource.Records.Count - 1
                    lngKey = Me.rptListSource.Records(intLoop).Item(mColP.����ID).Value
                    If chkRepeat(lngKey, Me.rptListSelect.Rows, CInt(mColP.����ID)) = False Then
                        Set Record = Me.rptListSelect.Records.Add
                        For intRow = 0 To Me.rptListSource.Columns.Count
                            Record.AddItem ""
                            Record.Item(intRow).Value = Me.rptListSource.Records(intLoop).Item(intRow).Value
                         Next
                    End If
                Next
                Me.rptListSource.Records.DeleteAll
            End If
            Me.rptListSelect.Populate: Me.rptListSource.Populate
        End If
    End If
    
    '��Ŀ
    If intType = 2 Then
        If intMove = 1 Then
            'Left
            If AllItem = False Then
                If Not Me.rptItemSelect.FocusedRow Is Nothing Then
                    lngKey = Me.rptItemSelect.FocusedRow.Record(mColI.ID).Value
                    If chkRepeat(lngKey, Me.rptItemSource.Rows, CInt(mColI.ID)) = False Then
                        Set Record = Me.rptItemSource.Records.Add
                        For intLoop = 0 To Me.rptItemSelect.Columns.Count
                            Record.AddItem ""
                            Record.Item(intLoop).Value = Me.rptItemSelect.FocusedRow.Record(intLoop).Value
                        Next
                        Me.rptItemSelect.Records.RemoveAt (Me.rptItemSelect.FocusedRow.Record.Index)
                    End If
                End If
            Else
                For intLoop = 0 To Me.rptItemSelect.Records.Count - 1
                    lngKey = Me.rptItemSelect.Records(intLoop).Item(mColI.ID).Value
                    If chkRepeat(lngKey, Me.rptItemSource.Rows, CInt(mColI.ID)) = False Then
                        Set Record = Me.rptItemSource.Records.Add
                        For intRow = 0 To Me.rptItemSelect.Columns.Count
                            Record.AddItem ""
                            Record.Item(intRow).Value = Me.rptItemSelect.Records(intLoop).Item(intRow).Value
                         Next
                    End If
                Next
                Me.rptItemSelect.Records.DeleteAll
            End If
            Me.rptItemSelect.Populate: Me.rptItemSource.Populate
        Else
            'Right
            If AllItem = False Then
                If Not Me.rptItemSource.FocusedRow Is Nothing Then
                    lngKey = Me.rptItemSource.FocusedRow.Record(mColI.ID).Value
                    If chkRepeat(lngKey, Me.rptItemSelect.Rows, CInt(mColI.ID)) = False Then
                        Set Record = Me.rptItemSelect.Records.Add
                        For intLoop = 0 To Me.rptItemSource.Columns.Count
                            Record.AddItem ""
                            Record.Item(intLoop).Value = Me.rptItemSource.FocusedRow.Record(intLoop).Value
                        Next
                        Me.rptItemSource.Records.RemoveAt (Me.rptItemSource.FocusedRow.Record.Index)
                        
                    End If
                End If
            Else
                For intLoop = 0 To Me.rptItemSource.Records.Count - 1
                    lngKey = Me.rptItemSource.Records(intLoop).Item(mColI.ID).Value
                    If chkRepeat(lngKey, Me.rptItemSelect.Rows, CInt(mColI.ID)) = False Then
                        Set Record = Me.rptItemSelect.Records.Add
                        For intRow = 0 To Me.rptItemSource.Columns.Count
                            Record.AddItem ""
                            Record.Item(intRow).Value = Me.rptItemSource.Records(intLoop).Item(intRow).Value
                         Next
                    End If
                Next
                Me.rptItemSource.Records.DeleteAll
            End If
            Me.rptItemSelect.Populate: Me.rptItemSource.Populate
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    zldatabase.SetPara "frmAddPatient_ѡ�����", Me.cbo����.ItemData(Me.cbo����.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_ѡ������", Me.cbo����.ItemData(Me.cbo����.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_ѡ�����", Me.cbo���.ItemData(Me.cbo���.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_��������", Me.cbo��������.ItemData(Me.cbo��������.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_����ҽ��", Me.cbo����ҽ��.ItemData(Me.cbo����ҽ��.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_ִ�п���", Me.cboִ�п���.ItemData(Me.cboִ�п���.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_��������", Me.cbo��������.ItemData(Me.cbo��������.ListIndex), 100, 1208
End Sub

Private Sub rptItemSelect_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call cmdItemLeft_Click
End Sub

Private Sub rptItemSource_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call cmdItemRight_Click
End Sub

Private Sub rptListSelect_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call cmdPatientLeft_Click
End Sub

Private Sub rptListSource_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call cmdPatientRight_Click
End Sub

Private Sub SaveData()
    '����                   ��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL() As String
    Dim intLoop As Integer, intItem As Integer
    Dim lngTmpID As Long, lngAdviceID As Long
    Dim lngMaxSeq As Long, iSendSeq   As Integer
    Dim intPatientType As Integer, lngPatientID As Long, intPatientPage As Integer
    Dim intBaby As Integer, lngExecDept As Long, lngPatientDept As Long, lngSendNO As Long
    Dim strDate As String, strDoctor As String, strNO As String, strAdviceText As String
    Dim strSample As String, strSampleNO As String, strName As String, strSex As String
    Dim lngCapID As Long, lngSampleID As Long, strlngID As String
    Dim strAge As String, strBed As String, strItemIDs As String, strItemResults As String
    Dim intMicrobe As Integer
    Dim blnJumpRepeat As Boolean
    Dim blnGetNO As Boolean
    Dim intNo As Integer
    Dim lngDeviceID As Long
    Dim strTmpDate As String
    Dim blnNew As Boolean
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "�������ɱ걾���Եȡ�����"
    
    ReDim strSQL(1 To 1)
    blnJumpRepeat = False
    
    With Me.cboִ�п���
        If .Text <> "" Then
            lngExecDept = .ItemData(.ListIndex)
        End If
    End With
    
    With Me.cbo��������
        If .Text <> "" Then
            lngPatientDetp = .ItemData(.ListIndex)
        End If
    End With
    
    With Me.cbo����ҽ��
        If .Text <> "" Then
            strDoctor = Mid(.Text, InStr(.Text, "-") + 1)
        End If
    End With
    
    With Me.cbo��������
        If .Text <> "" Then
            lngDeviceID = .ItemData(.ListIndex)
        End If
    End With
    
    For intLoop = 0 To Me.rptListSelect.Records.Count - 1
        
        '==============================================��������ҽ��==================================================
        strDate = "To_Date('" & Format(DTP, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        With Me.rptListSelect
            intPatientType = Val(.Records(intLoop).Item(mColP.������Դ).Value)
            intPatientType = IIf(intPatientType = 0, 3, intPatientType)
            lngPatientID = Val(.Records(intLoop).Item(mColP.����ID).Value)
            intPatientPage = Val(.Records(intLoop).Item(mColP.��ҳID).Value)
            intBaby = Val(.Records(intLoop).Item(mColP.Ӥ��).Value)
            strName = .Records(intLoop).Item(mColP.����).Value
            strSex = .Records(intLoop).Item(mColP.�Ա�).Value
            strAge = .Records(intLoop).Item(mColP.����).Value
            strlngID = Val(.Records(intLoop).Item(mColP.��ʶ��).Value)
            strBed = .Records(intLoop).Item(mColP.����).Value
        End With

        lngAdviceID = zldatabase.GetNextId("����ҽ����¼")              '���ID
        lngSendNO = zldatabase.GetNextNo(10)                            '���ͺ�
        strNO = zldatabase.GetNextNo(IIf(PatientType = 2, 14, 13))      '���ݺ�
        
        '=======ȡ�ĵ��ݺŻ��ظ�����֪��ԭ���ȴ���Ϊ��������ظ�����ȡһ�Ρ���ʱ����
        intNo = 0
        blnGetNO = False
        Do Until blnGetNO = False
            intNo = intNo + 1
            gstrSql = "select " & gConst_����ҽ������_���� & " from ����ҽ������ a where no = [1] "
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strNO)
            If rsTmp.EOF = False Then blnGetNO = True
            If intNo >= 10 Then blnGetNO = True
        Loop
        '======================================================================
        '�õ����ҽ�����
        gstrSql = "select max(���) as ��� from ����ҽ����¼ where ����id = [1] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.ControlBox, lngPatientID)
        If rsTmp.EOF = False Then
            lngMaxSeq = Val(Nvl(rsTmp("���"), 0))
        Else
            lngMaxSeq = 0
        End If

        iSendSeq = 1
        For intItem = 0 To Me.rptItemSelect.Records.Count - 1
            '������Ŀҽ��
            lngTmpID = zldatabase.GetNextId("����ҽ����¼")
            lngMaxSeq = lngMaxSeq + 1
            If intItem = 0 Then
                strAdviceText = Replace(rptItemSelect.Records(intItem).Item(mColI.����).Value, "'", "''") & "(" & _
                    rptItemSelect.Records(intItem).Item(mColI.�걾).Value & ")"
                strSample = rptItemSelect.Records(intItem).Item(mColI.�걾).Value
                intMicrobe = rptItemSelect.Records(intItem).Item(mColI.��Ŀ���).Value
            Else
                strAdviceText = Replace(rptItemSelect.Records(intItem).Item(mColI.����).Value, "'", "''") & "," & strAdviceText
            End If
            strItemIDs = strItemIDs & "," & rptItemSelect.Records(intItem).Item(mColI.ID).Value
            strSQL(ReDimArray(strSQL)) = "ZL_����ҽ����¼_Insert(" & lngTmpID & "," & lngAdviceID & "," & lngMaxSeq & "," & intPatientType & _
                "," & lngPatientID & "," & IIf(intPatientPage = 0, "NULL", intPatientPage) & "," & IIf(intBaby = 0, "NULL", intBaby) & _
                ",1,1,'C'," & rptItemSelect.Records(intItem).Item(mColI.ID).Value & ",NULL,NULL,NULL,NULL,'" & _
                Replace(rptItemSelect.Records(intItem).Item(mColI.����).Value, "'", "''") & "',NULL,'" & _
                rptItemSelect.Records(intItem).Item(mColI.�걾).Value & "','һ����',NULL,NULL,NULL,NULL,0," & lngExecDept & ",4,0," & _
                strDate & ",NULL," & lngPatientDetp & "," & lngPatientDetp & ",'" & strDoctor & "'," & strDate & ")"

            iSendSeq = iSendSeq + 1
            strSQL(ReDimArray(strSQL)) = "ZL_����ҽ������_Insert(" & lngTmpID & "," & lngSendNO & "," & IIf(intPatientType = 2, 2, 1) & _
            ",'" & strNO & "'," & iSendSeq & ",NULL,NULL,NULL,Sysdate+1/(24*3600),0," & lngExecDept & ",0,0)"
        Next

        '�ɼ���ʽҽ��
        gstrSql = "Select �÷�id From ������ĿĿ¼ A, �����÷����� B Where A.ID = B.��Ŀid and a.id = [1] and b.���� = 1 "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.rptItemSelect.Records(0).Item(mColI.ID).Value))
        'û���ҵ��ɼ���ʽʱ�˳�
        If rsTmp.EOF = True Then MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName: Exit Sub
        lngCapID = rsTmp("�÷�ID")
        lngMaxSeq = lngMaxSeq + 1

        strSQL(ReDimArray(strSQL)) = "ZL_����ҽ����¼_Insert(" & lngAdviceID & ",NULL," & lngMaxSeq & "," & intPatientType & _
                "," & lngPatientID & "," & IIf(intPatientPage = 0, "NULL", intPatientPage) & "," & IIf(intBaby = 0, "NULL", intBaby) & _
                ",1,1,'E'," & lngCapID & ",NULL,NULL,NULL,NULL,'" & _
                strAdviceText & "',NULL,'" & _
                strSample & "','һ����',NULL,NULL,NULL,NULL,2," & lngExecDept & ",3,0," & _
                strDate & ",NULL," & lngPatientDetp & "," & lngPatientDetp & ",'" & strDoctor & "'," & strDate & ")"

        iSendSeq = iSendSeq + 1
        strSQL(ReDimArray(strSQL)) = "ZL_����ҽ������_Insert(" & lngAdviceID & "," & lngSendNO & "," & IIf(intPatientType = 2, 2, 1) & _
            ",'" & strNO & "'," & iSendSeq & ",NULL,NULL,NULL,Sysdate+1/(24*3600),0," & lngExecDept & ",0,1)"
        '====================================================================================================================================

        '=================================================================�걾��Ϣ===========================================================
        '�걾��
        strSampleNO = Val(Me.txt���ɱ걾��.Text) + intLoop
        If lngDeviceID = -1 Then
            strSampleNO = TransSampleNO(Val(txt�걾����.Text) & "-" & Val(strSampleNO))
        End If
        
        gstrSql = "Select Id,�걾���" & vbNewLine & _
                    "From ����걾��¼" & vbNewLine & _
                    "Where ����ʱ�� Between To_Date(To_Char([1], 'yyyy-MM-dd') || ' 00:00:00', 'yyyy-MM-dd HH24:mi:ss') And" & vbNewLine & _
                    "           To_Date(To_Char([1], 'yyyy-MM-dd') || ' 23:59:59', 'yyyy-MM-dd HH24:mi:ss') And �걾��� = [2] And" & vbNewLine & _
                    "           Nvl(����id, 0) = Nvl([3], 0) " & IIf(blnEmergency = 1, " and nvl(�걾����,0) = [4]", "")
        strTmpDate = Format(zldatabase.Currentdate, "yyyy-mm-dd")
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(strTmpDate), strSampleNO, _
        Val(IIf(lngDeviceID = -1, 0, lngDeviceID)), 0)
        If rsTmp.EOF = False Then
            If blnJumpRepeat = False Then
                If MsgBox("�����б걾���ظ��Ƿ����̣�", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
                blnJumpRepeat = True
            End If
            lngSampleID = rsTmp("ID")
            blnNew = False
        Else
            lngSampleID = zldatabase.GetNextId("����걾��¼")
            blnNew = True
        End If
        
        
        strSQL(ReDimArray(strSQL)) = "ZL_����걾��¼_�걾����(" & lngSampleID & "," & lngAdviceID & ",'" & lngAdviceID & "',0,'" & strSampleNO & "'," & _
            strDate & ",NULL," & IIf(lngDeviceID = -1, "NULL", lngDeviceID) & "," & strDate & ",NULL,'" & UserInfo.���� & "'," & _
            strDate & "," & IIf(intMicrobe = 2, 1, "Null") & ",0,NULL,'" & strName & "','" & strSex & "','" & strAge & "','" & lngSendNO & "','" & strSample & "'," & _
            lngPatientDetp & ",'" & strDoctor & "'," & IIf(strlngID = 0, "NULL", strlngID) & ",'" & strBed & "'," & lngPatientDetp & ",'" & _
            strAdviceText & "',1," & lngPatientID & "," & cboִ�п���.ItemData(cboִ�п���.ListIndex) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            
        If blnNew = True Then
            If intMicrobe = 2 Then
                gstrSql = "Select Id, ԭʼ���, ������, �����־, ����ο�, Rownum As �������, ������Ŀid" & vbNewLine & _
                        "From (Select d.ϸ��id As Id, '' As ԭʼ���, c.Ĭ�Ͻ�� As ������, '' As �����־, '' As ����ο�, d.�������," & vbNewLine & _
                        "                           a.Id As ������Ŀid" & vbNewLine & _
                        "            From ������ĿĿ¼ a, ���鱨����Ŀ d, ����ϸ�� c" & vbNewLine & _
                        "            Where a.Id = d.������Ŀid And d.ϸ��id = c.Id And a.Id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & vbNewLine & _
                        "            Order By a.����, d.�������)"
            Else
                gstrSql = "Select Id, ԭʼ���, ������, �����־, Rownum As �������, ������Ŀid,����ο�" & vbNewLine & _
                            "From (Select d.������Ŀid As Id, '' As ԭʼ���, Decode(c.�������, 3, Nvl(c.Ĭ��ֵ, '-'), 2, c.Ĭ��ֵ, '') As ������," & vbNewLine & _
                            "                           '' As �����־, d.�������, a.Id As ������Ŀid," & vbNewLine & _
                            "                           Trim(Replace(Replace(' ' || Zlgetreference(d.������Ŀid, a.�걾��λ, 0, Null), ' .', '0.'), '��.', '��0.')) As ����ο�" & vbNewLine & _
                            "            From ������ĿĿ¼ a, ���鱨����Ŀ d, ������Ŀ c" & vbNewLine & _
                            "            Where a.Id = d.������Ŀid And d.������Ŀid = c.������Ŀid And d.ϸ��id Is Null And c.��Ŀ��� <> 2 And" & vbNewLine & _
                            "                        a.Id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & vbNewLine & _
                            "            Order By a.����, d.�������) "
            End If
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(strItemIDs, 2))
            strItemResults = ""
            Do While Not rsTmp.EOF
                strItemResults = strItemResults & "|" & lngAdviceID & "^" & Nvl(rsTmp("ID")) & "^" & Nvl(rsTmp("������")) & _
                    "^" & Nvl(rsTmp("�����־")) & "^" & Nvl(rsTmp("����ο�")) & "^" & Nvl(rsTmp("������ĿID")) & _
                    "^" & Nvl(rsTmp("�������"))
                rsTmp.MoveNext
            Loop
            
            strSQL(ReDimArray(strSQL)) = "Zl_������ͨ���_Write(" & lngSampleID & "," & IIf(lngDeviceID = -1, "NULL", lngDeviceID) & ",'" & _
                Mid(strItemResults, 2) & "',0," & IIf(intMicrobe = 2, 1, 0) & ")"
            strSQL(ReDimArray(strSQL)) = "Zl_���¼�����_Cale(" & lngSampleID & ")"
        End If
        '====================================================================================================================================
    Next
    '��ʼִ��
    On Error GoTo errH
'    gcnOracle.BeginTrans
    
    
    For intLoop = 1 To UBound(strSQL)
        Debug.Print strSQL(intLoop) & vbCrLf
        If strSQL(intLoop) <> "" Then zldatabase.ExecuteProcedure strSQL(intLoop), Me.Caption
    Next
    zlCommFun.StopFlash
'    gcnOracle.CommitTrans
    MsgBox "���ɱ걾���!", vbInformation, Me.Caption
    Me.MousePointer = 0
    Exit Sub
errH:
    zlCommFun.StopFlash
    Me.MousePointer = 0
'    gcnOracle.RollbackTrans
    If errcenter() = 1 Then Resume
    Call saveerrlog
End Sub
  
Private Sub txt�걾��_GotFocus()
    txt�걾��.SelStart = 0
    txt�걾��.SelLength = Len(txt�걾��)
End Sub

Private Sub txt�걾��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmd����_Click: Me.txt�걾��.SelStart = 0: Me.txt�걾��.SelLength = Len(Me.txt�걾��)
    If InStr("1234567890~,-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt�걾����_GotFocus()
    txt�걾����.SelStart = 0
    txt�걾����.SelLength = Len(txt�걾����)
End Sub

Private Sub txt�걾����_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0
    Me.txt����.SelLength = Len(Me.txt����.Text)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmd����_Click
    End If
End Sub
Private Function chkSaveData() As Boolean
    '����           ������
    
    If rptListSelect.Records.Count = 0 Then
        MsgBox "������ϢΪ�գ���ѡ������Ϣ!", vbInformation, Me.Caption
        Exit Function
    End If
    
    If rptItemSelect.Records.Count = 0 Then
        MsgBox "��Ŀ��ϢΪ�գ���ѡ����Ŀ��Ϣ!", vbInformation, Me.Caption
        Exit Function
    End If
    
    If cbo��������.Text = "" Then
        MsgBox "��ѡ�񿪵�����!", vbInformation, Me.Caption
        Me.cbo��������.SetFocus
        Exit Function
    End If
    
    If cbo����ҽ��.Text = "" Then
        MsgBox "��ѡ�񿪵�ҽ��!", vbInformation, Me.Caption
        Me.cbo����ҽ��.SetFocus
        Exit Function
    End If
    
    If cboִ�п���.Text = "" Then
        MsgBox "��ѡ��ִ�п���!", vbInformation, Me.Caption
        Me.cboִ�п���.SetFocus
        Exit Function
    End If
    
    If cbo��������.Text = "" Then
        MsgBox "��ѡ���������!", vbInformation, Me.Caption
        Me.cbo��������.SetFocus
        Exit Function
    End If
    
    If cbo��������.ItemData(cbo��������.ListIndex) = -1 Then
        If Trim(txt�걾����) = "" Or Trim(txt���ɱ걾��) = "" Then
            MsgBox "�ֹ���Ŀ�����������źͱ걾��!", vbInformation, Me.Caption
            Me.txt����.SetFocus
            Exit Function
        End If
    Else
        If Trim(txt���ɱ걾��) = "" Then
            MsgBox "���������ɵı걾�ſ�ʼ��!", vbInformation, Me.Caption
            Me.txt���ɱ걾��.SetFocus
            Exit Function
        End If
    End If
    
    chkSaveData = True
    
    
End Function

Private Sub txt����_GotFocus()
    txt����.SelStart = 0
    txt����.SelLength = Len(txt����.Text)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt���ɱ걾��_GotFocus()
    txt���ɱ걾��.SelStart = 0
    txt���ɱ걾��.SelLength = Len(txt���ɱ걾��)
End Sub

Private Sub txt���ɱ걾��_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub
