VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDiffPriceRecalCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩƷ��ۼ���"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmDiffPriceRecalCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdInput 
      Caption         =   "¼����(&E)"
      Height          =   350
      Left            =   3525
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   285
      Left            =   1245
      TabIndex        =   20
      Top             =   3270
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
      Format          =   114491395
      CurrentDate     =   36444
      MaxDate         =   401768
   End
   Begin VB.CommandButton cmdIni 
      Caption         =   "��ʼ���(&I)"
      Height          =   350
      Left            =   45
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "�鿴(&B)"
      Height          =   350
      Left            =   2430
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "��˽��(&V)"
      Height          =   350
      Left            =   45
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ȡ�����(&D)"
      Height          =   350
      Left            =   1245
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ComboBox cbo���㷽�� 
      Height          =   300
      Left            =   5250
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2895
      Width           =   1605
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7065
      TabIndex        =   5
      Top             =   3840
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   30
      Left            =   -840
      TabIndex        =   4
      Top             =   3615
      Width           =   9060
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5970
      TabIndex        =   2
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "����(&O)"
      Height          =   350
      Left            =   4860
      TabIndex        =   1
      Top             =   3840
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4395
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiffPriceRecalCard.frx":000C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9816
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cbo�ⷿ 
      Height          =   300
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2910
      Width           =   1605
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   285
      Left            =   5250
      TabIndex        =   16
      Top             =   3255
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
      Format          =   114491395
      CurrentDate     =   36444
      MaxDate         =   401768
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   285
      Left            =   5250
      TabIndex        =   21
      Top             =   3255
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
      Format          =   114491395
      CurrentDate     =   36444
      MaxDate         =   401768
   End
   Begin VB.Label lblEnd 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��"
      Height          =   180
      Left            =   4290
      TabIndex        =   19
      Top             =   3315
      Width           =   720
   End
   Begin VB.Label lblBegin 
      AutoSize        =   -1  'True
      Caption         =   "��ʼʱ��"
      Height          =   180
      Left            =   420
      TabIndex        =   18
      Top             =   3345
      Width           =   720
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "���ν��ʱ��"
      Height          =   180
      Left            =   4125
      TabIndex        =   15
      Top             =   3315
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lbl�ϴν�� 
      Caption         =   "2007-01-01 22:00:00(δ���)"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1245
      TabIndex        =   11
      Top             =   3345
      Visible         =   0   'False
      Width           =   6660
   End
   Begin VB.Label lbl��� 
      AutoSize        =   -1  'True
      Caption         =   "�ϴν��ʱ��"
      Height          =   180
      Left            =   75
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���㷽��"
      Height          =   180
      Left            =   4290
      TabIndex        =   8
      Top             =   2955
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�ⷿ"
      Height          =   180
      Left            =   450
      TabIndex        =   6
      Top             =   2985
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmDiffPriceRecalCard.frx":08A0
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblMemo 
      Caption         =   $"frmDiffPriceRecalCard.frx":0CE2
      ForeColor       =   &H00C00000&
      Height          =   2505
      Left            =   630
      TabIndex        =   3
      Top             =   75
      Width           =   7620
   End
End
Attribute VB_Name = "frmDiffPriceRecalCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr�ϴν��ʱ�� As String
Dim mbln�Ƿ���� As Boolean
Dim mlng�ⷿID As Long
Dim mint���㷽�� As Integer
Dim mint��λϵ�� As Integer
Dim mbln����ʼ��� As Boolean

Private Const intIni As Integer = 6

Private Enum con���㷽��
    type_�ƶ�ƽ�� = 1
    type_ȫ��ƽ�� = 2
    type_�Ƚ��ȳ� = 3
End Enum
Private Sub GetUnit(ByVal lng�ⷿID As Long)
    Dim strUnit As String
    strUnit = GetDrugUnit(lng�ⷿID, Me.Caption)
    Select Case strUnit
        Case "סԺ��λ"
            mint��λϵ�� = 4
        Case "���ﵥλ"
            mint��λϵ�� = 3
        Case "ҩ�ⵥλ"
            mint��λϵ�� = 2
        Case "�ۼ۵�λ"
            mint��λϵ�� = 1
    End Select
End Sub

Private Sub Get�ϴν��(ByVal lng�ⷿID As Long)
    Dim rsTmp As New ADODB.Recordset
    
    mbln����ʼ��� = False
    
    On Error GoTo errHandle
    '���ѡ���Ƚ��ȳ�������ʾ�ϴν����Ϣ
    If cbo���㷽��.ListIndex = 1 Then
        lblTime.Visible = True
        dtpTime.Visible = True
        dtpTime.Value = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        '����Ƿ�����ڳ�ʼ���
        gstrSQL = "Select nvl(�Ƿ��ʼ,0) �Ƿ��ʼ From ҩƷ��� Where Nvl(�Ƿ��ʼ, 0) = 1 And �ⷿid = [1] And Rownum = 1" & _
                " Union All " & _
                " Select nvl(�Ƿ��ʼ,0) �Ƿ��ʼ From ҩƷ��� Where Nvl(�Ƿ��ʼ, 0) = 0 And �ⷿid = [1] And Rownum = 1"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-����Ƿ���ڳ�ʼ���", lng�ⷿID)
        If rsTmp.RecordCount = 1 Then
            If rsTmp!�Ƿ��ʼ = 1 Then
                mbln����ʼ��� = True
            End If
        End If
        cmdInput.Visible = mbln����ʼ���
        
        '����Ƿ�����ϴν��
        gstrSQL = "Select Max(�������) ���ʱ�� From ҩƷ���  Where �ⷿid=[1] "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-����Ƿ�����ϴν��", lng�ⷿID)
        
        mstr�ϴν��ʱ�� = Format(rsTmp!���ʱ��, "YYYY-MM-DD HH:MM:SS")
        
        If mstr�ϴν��ʱ�� = "" Then
            lbl���.Visible = True
            lbl�ϴν��.Visible = True
            cmdIni.Visible = True
            lbl�ϴν��.Caption = "�޳�ʼ�����Ϣ�����ʼ����棡"
            Exit Sub
        End If
        
        'ȡ�ϴν����Ϣ
        gstrSQL = "Select Nvl(����־, 0) ����־ From ҩƷ��� " & _
             " Where �ⷿid = [1] And ������� = [2] And Rownum = 1"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-ȡ�ϴν��", lng�ⷿID, CDate(Format(mstr�ϴν��ʱ��, "yyyy-mm-dd hh:mm:ss")))
        
        If rsTmp.RecordCount = 0 Then
            mstr�ϴν��ʱ�� = ""
            lbl���.Visible = True
            lbl�ϴν��.Visible = True
            cmdIni.Visible = True
            lbl�ϴν��.Caption = "�޳�ʼ�����Ϣ�����ʼ����棡"
            Exit Sub
        Else
            mbln�Ƿ���� = (rsTmp!����־ = 1)
        End If
        
        '��ʾ�ϴν����Ϣ
        If mstr�ϴν��ʱ�� <> "" Then
            lbl���.Visible = True
            lbl�ϴν��.Visible = True
            cmdVerify.Visible = True
            cmdDel.Visible = True
            cmdBrowse.Visible = True
            lbl�ϴν��.Caption = mstr�ϴν��ʱ�� & IIf(mbln�Ƿ����, "", "(δ���)")
            cmdVerify.Enabled = Not mbln�Ƿ����
        End If
        
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub IniControl()
    lbl���.Visible = False
    lbl�ϴν��.Visible = False
    cmdVerify.Visible = False
    cmdDel.Visible = False
    cmdBrowse.Visible = False
    lblTime.Visible = False
    dtpTime.Visible = False
    cmdIni.Visible = False
    
    lblBegin.Visible = False
    lblEnd.Visible = False
    dtpBegin.Visible = False
    dtpEnd.Visible = False
End Sub

Private Sub RefreshNow(ByVal lng�ⷿID As Long)
    Call IniControl
    
    If mint���㷽�� = type_ȫ��ƽ�� Then
        lblBegin.Visible = True
        lblEnd.Visible = True
        dtpBegin.Visible = True
        dtpEnd.Visible = True
        
        dtpBegin.Value = Format(Sys.Currentdate, "yyyy-mm") & "-01 00:00:00"
        dtpEnd.Value = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        Exit Sub
    End If
    
    If mint���㷽�� = type_�Ƚ��ȳ� Then
        Call Get�ϴν��(lng�ⷿID)
    End If
End Sub

Private Sub cbo���㷽��_Click()
    Select Case cbo���㷽��.ListIndex
        Case 0
            'ȫ��ƽ��
            mint���㷽�� = type_ȫ��ƽ��
        Case 1
            '�Ƚ��ȳ�
            mint���㷽�� = type_�Ƚ��ȳ�
    End Select
    
    Call RefreshNow(mlng�ⷿID)
End Sub


Private Sub cbo�ⷿ_Click()
    If Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex) <> mlng�ⷿID Then
        mlng�ⷿID = Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex)
        Call GetUnit(mlng�ⷿID)
        Call RefreshNow(mlng�ⷿID)
    End If
End Sub


Private Sub cmdBrowse_Click()
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1308", Me, "�ⷿ=" & Cbo�ⷿ.Text & "|" & IIf(Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex) = 0, " is not null ", "=" & Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex)), "�������=" & CDate(mstr�ϴν��ʱ��), "��λ=" & mint��λϵ��)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    If MsgBox("�Ƿ�ɾ���ϴν�棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    gstrSQL = "Zl_ҩƷ���_Delete(to_date('" & Format(mstr�ϴν��ʱ��, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')," & mlng�ⷿID & ",2)"
    Me.staThis.Panels(2).Text = "����ɾ���ϴν�棬��ȴ���������"
    
    Me.MousePointer = vbHourglass
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Me.MousePointer = vbDefault
    MsgBox "ɾ���ɹ���", vbOKOnly + vbInformation, gstrSysName
    
    Call IniControl
    Call Get�ϴν��(mlng�ⷿID)
    
    DoEvents
    Me.staThis.Panels(2).Text = ""
    
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdIni_Click()
    '��������Ƚ��ȳ�������δ�г�ʼ���棬��Ҫ���ɳ�ʼ���
    If mstr�ϴν��ʱ�� = "" And DateDiff("m", dtpTime.Value, Sys.Currentdate) > intIni Then
        MsgBox "�ڳ�������ڲ�������" & intIni & "���¡�"
        dtpTime.Value = Sys.Currentdate
        Exit Sub
    End If
    
    If mint���㷽�� = type_�Ƚ��ȳ� And mstr�ϴν��ʱ�� = "" Then
        gstrSQL = "Zl_ҩƷ���_Insert(to_date('" & Format(dtpTime.Value, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss') ," & mlng�ⷿID & ",NULL)"
    Else
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    Me.staThis.Panels(2).Text = "���ڳ�ʼҩƷ��棬��ȴ���������"
    
    Me.MousePointer = vbHourglass
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Me.MousePointer = vbDefault
    MsgBox "��ʼҩƷ���ɹ���", vbOKOnly + vbInformation, gstrSysName
    
    Call IniControl
    Call Get�ϴν��(mlng�ⷿID)
    
    DoEvents
    Me.staThis.Panels(2).Text = ""
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdInput_Click()
    If MsgBox("¼���۱������пⷿ����ɳ�ʼ��棬�Ƿ����ھ�¼���ۣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    frmDiffPriceRecal.ShowCard Me, 2
    
End Sub
Private Sub CmdSave_Click()
    Dim strFirstTime As String
    Dim strSecondTime As String
    
    On Error GoTo errHandle
    
    If mint���㷽�� = type_�Ƚ��ȳ� And mstr�ϴν��ʱ�� = "" Then
        MsgBox "�ÿⷿ��û�н��г�ʼ��棬�밴��ʼ��水ť���н�棡", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mint���㷽�� = type_�Ƚ��ȳ� And mstr�ϴν��ʱ�� <> "" And mbln�Ƿ���� = False Then
        MsgBox "�ÿⷿ�ϴν�滹û����ˣ���������ϴν�棡", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mint���㷽�� = type_�Ƚ��ȳ� And mstr�ϴν��ʱ�� <> "" And mbln�Ƿ���� = True Then
        If mstr�ϴν��ʱ�� > Format(dtpTime.Value, "yyyy-mm-dd hh:mm:ss") Then
            MsgBox "��ǰ�Ľ������С�����ϴν�����ڣ����������ý�����ڣ�����ȡ���ϴν�棡", vbOKOnly + vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If mint���㷽�� = type_�Ƚ��ȳ� Then
        strFirstTime = Format(dtpTime.Value, "yyyy-mm-dd hh:mm:ss")
        strSecondTime = mstr�ϴν��ʱ��
    End If
    
    If mint���㷽�� = type_ȫ��ƽ�� Then
        strFirstTime = Format(dtpBegin.Value, "yyyy-mm-dd hh:mm:ss")
        strSecondTime = Format(dtpEnd.Value, "yyyy-mm-dd hh:mm:ss")
    End If
    
    gstrSQL = "zl_ҩƷ�������_UPDATE(to_date('" & Format(strFirstTime, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss') ," & mlng�ⷿID & ","

    If mstr�ϴν��ʱ�� = "" Then
        gstrSQL = gstrSQL & "NULL,"
    Else
        gstrSQL = gstrSQL & "to_date('" & Format(strSecondTime, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    End If

    gstrSQL = gstrSQL & mint���㷽�� & ")"
      
    Me.staThis.Panels(2).Text = "���ڼ����ۣ���ȴ���������"
    
    Me.MousePointer = vbHourglass
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Me.MousePointer = vbDefault
    MsgBox "�������ɹ���", vbOKOnly + vbInformation, gstrSysName
    
    Call RefreshNow(mlng�ⷿID)
    
    DoEvents
    Me.staThis.Panels(2).Text = ""
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdVerify_Click()
    gstrSQL = "Zl_ҩƷ���_Verify(to_date('" & Format(mstr�ϴν��ʱ��, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')," & mlng�ⷿID & ")"
    Me.staThis.Panels(2).Text = "��������ϴν�棬��ȴ���������"
    
    Me.MousePointer = vbHourglass
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Me.MousePointer = vbDefault
    MsgBox "��˳ɹ���", vbOKOnly + vbInformation, gstrSysName

    Call IniControl
    Call Get�ϴν��(mlng�ⷿID)
    
    DoEvents
    Me.staThis.Panels(2).Text = ""
    
    Exit Sub
End Sub

Private Sub dtpBegin_Change()
    If DateDiff("s", dtpEnd.Value, dtpBegin.Value) > 0 Then
        dtpBegin.Value = Format(Sys.Currentdate, "yyyy-mm") & "-01 00:00:00"
    End If
End Sub


Private Sub dtpEnd_Change()
    If DateDiff("s", dtpEnd.Value, dtpBegin.Value) > 0 Then
        dtpEnd.Value = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New Recordset
    
    On Error GoTo errHandle
    RestoreWinState Me, App.Title
    
    '����ⷿ
    gstrSQL = "Select Distinct A.ID, A.���� " & _
            " From ��������˵�� C, �������ʷ��� B, ���ű� A " & _
            " Where (a.վ�� = [1] Or a.վ�� is Null) And C.�������� = B.���� And A.ID = C.����id And " & _
            " To_Char(A.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' And " & _
            " Instr('HIJKLMN', B.����, 1) > 0"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-�������пⷿ", gstrNodeNo)
    
    With Cbo�ⷿ
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!����
            .ItemData(.NewIndex) = rsTmp!id
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        'If .ListIndex = -1 Then .ListIndex = 0
        If .ListCount < 1 Then
            MsgBox "����Ӧ������һ����ҩ�����ʣ�ҩ�����ʣ������Ƽ������ʵĲ��ţ���鿴���Ź���", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        Else
            .ListIndex = 0
        End If
    End With
    
    '������㷽��
    With cbo���㷽��
        .Clear
        .AddItem "ȫ��ƽ��"
        .AddItem "�Ƚ��ȳ�"
        If .ListIndex = -1 Then .ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.Title
End Sub

