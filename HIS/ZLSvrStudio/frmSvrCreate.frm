VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSvrCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������"
   ClientHeight    =   5115
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "frmSvrCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7950
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraTbs 
      Height          =   1035
      Left            =   2310
      TabIndex        =   4
      Top             =   1380
      Width           =   5145
      Begin VB.TextBox txtTbsFile 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   630
         TabIndex        =   6
         Top             =   600
         Width           =   4185
      End
      Begin VB.TextBox txtTmpFile 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   630
         TabIndex        =   7
         Top             =   600
         Width           =   4185
      End
      Begin VB.TextBox txtTbsSize 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4050
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "200"
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtTmpSize 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4050
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "50"
         Top             =   180
         Width           =   555
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   225
         TabIndex        =   11
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblTbsFile 
         AutoSize        =   -1  'True
         Caption         =   "�ļ�"
         Height          =   180
         Left            =   225
         TabIndex        =   10
         Top             =   660
         Width           =   360
      End
      Begin VB.Label lblTbsSize 
         AutoSize        =   -1  'True
         Caption         =   "��С        M"
         Height          =   180
         Left            =   3645
         TabIndex        =   9
         Top             =   255
         Width           =   1170
      End
      Begin VB.Label lblTbsName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "zlToolsTbs"
         Height          =   300
         Left            =   630
         TabIndex        =   14
         Top             =   210
         Width           =   1590
      End
      Begin VB.Label lblTmpName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "zlToolsTmp"
         Height          =   300
         Left            =   630
         TabIndex        =   15
         Top             =   210
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdRegFile 
      Caption         =   "ѡ��(&R)��"
      Height          =   350
      Left            =   6585
      TabIndex        =   26
      Top             =   3330
      Width           =   1100
   End
   Begin VB.TextBox txtPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2205
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "ZLSOFT"
      Top             =   360
      Width           =   2160
   End
   Begin MSComctlLib.ProgressBar pgbState 
      Height          =   150
      Left            =   2595
      TabIndex        =   20
      Top             =   4875
      Visible         =   0   'False
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   1995
      TabIndex        =   18
      Top             =   4245
      Width           =   1100
   End
   Begin VB.CommandButton cmdSqlFile 
      Caption         =   "ѡ��(&S)��"
      Height          =   350
      Left            =   6570
      TabIndex        =   12
      Top             =   2625
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6585
      TabIndex        =   2
      Top             =   4245
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5385
      TabIndex        =   1
      Top             =   4245
      Width           =   1100
   End
   Begin VB.PictureBox PicSetup 
      Align           =   3  'Align Left
      Height          =   4740
      Left            =   0
      ScaleHeight     =   4680
      ScaleWidth      =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   1740
      Begin VB.FileListBox fltFile 
         Appearance      =   0  'Flat
         Height          =   1104
         Left            =   135
         Pattern         =   "*.zcr"
         TabIndex        =   27
         Top             =   3405
         Visible         =   0   'False
         Width           =   1410
      End
      Begin MSComDlg.CommonDialog dlgMain 
         Left            =   255
         Top             =   2835
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgSetup 
         Height          =   2625
         Left            =   60
         Picture         =   "frmSvrCreate.frx":058A
         Stretch         =   -1  'True
         Top             =   -105
         Width           =   945
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   19
      Top             =   4740
      Width           =   7944
      _ExtentX        =   14023
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1746
            MinWidth        =   882
            Text            =   "��װ���� "
            TextSave        =   "��װ���� "
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8493
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "16:59"
            Key             =   "STANUM"
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
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   855
      TabIndex        =   21
      Top             =   4095
      Width           =   6975
   End
   Begin MSComctlLib.TabStrip tbsTbs 
      Height          =   1440
      Left            =   2205
      TabIndex        =   28
      ToolTipText     =   "���ع���,�Զ��������ߴ�(AUTOALLOCATE),�������ʱ��ռ�,��ͳһ���ߴ�1M"
      Top             =   1080
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   2540
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ȱʡ�ռ�"
            Key             =   "Tbs"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ʱ�ռ�"
            Key             =   "Tmp"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label lblRegFile 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2190
      TabIndex        =   25
      Top             =   3645
      Width           =   5490
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "4)ϵͳ������Ҫע����Ȩ�ļ����밴ָ��ѡ��"
      Height          =   180
      Index           =   3
      Left            =   1995
      TabIndex        =   24
      Top             =   3375
      Width           =   3780
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "(Ĭ������Ϊ""ZLSOFT"")"
      Height          =   180
      Left            =   4410
      TabIndex        =   23
      Top             =   435
      Width           =   1800
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "1)�������������û��̶�Ϊ""zlTools""�������������������룺"
      Height          =   180
      Index           =   0
      Left            =   1995
      TabIndex        =   22
      Top             =   90
      Width           =   5130
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "3)�����ߴ��������ڽű��ļ����밴ָ��ѡ��"
      Height          =   180
      Index           =   2
      Left            =   1995
      TabIndex        =   17
      Top             =   2670
      Width           =   3960
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "2)��������洢һ���������ݣ���ȷ�����ռ��λ�����С��"
      Height          =   180
      Index           =   1
      Left            =   1995
      TabIndex        =   16
      Top             =   840
      Width           =   5220
   End
   Begin VB.Label lblSqlFile 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2190
      TabIndex        =   13
      Top             =   2940
      Width           =   5475
   End
End
Attribute VB_Name = "frmSvrCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrTbsPath As String                        'ȱʡ��ռ�·�����ƣ�������ʷ��ռ����
Private mobjFiles As New FileSystemObject
Private mobjText As TextStream

Private cnTools As New ADODB.Connection

Private mclsRunScript As clsRunScript
'��ʱ����
Dim rsTemp As New ADODB.Recordset
Dim strSQL As String, strTemp As String
Dim lngCount As Long

Private Sub cmdCancel_Click()
    If MsgBox("��δ���������ߣ����ȡ����", vbQuestion + vbYesNo, "��ʾ") = vbNo Then Exit Sub
    Unload Me
End Sub

Private Sub cmdRegFile_Click()
    With Me.dlgMain
        .FileName = lblRegFile.Caption
        .DialogTitle = "ѡ��ע����Ȩ�ļ�"
        .Filter = "(ע����Ȩ�ļ�)|*.zcr"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            lblRegFile.Caption = .FileName
        End If
    End With
End Sub

Private Sub cmdSqlFile_Click()
    With Me.dlgMain
        .FileName = lblSqlFile.Caption
        .DialogTitle = "ѡ������߽ű��ļ�"
        .Filter = "(�����߽ű��ļ�)|zlServer.sql;*.plb"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            Me.fltFile.Path = Mid(.FileName, 1, Len(.FileName) - InStr(1, StrReverse(.FileName), "\") + 1)
            Me.fltFile.Pattern = IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "zlRegist.plb")
            If Me.fltFile.ListCount = 0 Then
                lblSqlFile.Caption = ""
                MsgBox "��λ��δ������Ȩ��֤�ļ���", vbExclamation, gstrSysName
            Else
                lblSqlFile.Caption = .FileName
            End If
        End If
    End With
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub cmdOK_Click()
    
    If Trim(Me.txtPwd.Text) = "" Then
        MsgBox "û�����ù��������������룬���ܼ�����", vbExclamation, "��ʾ"
        Me.txtPwd.SetFocus
        Exit Sub
    End If

    If Val(txtTbsSize.Text) < 100 Then
        MsgBox "ȱʡ�ռ��С���ô���", vbExclamation, "��ʾ"
        txtTbsSize.Text = 100
        If txtTbsSize.Visible Then txtTbsSize.SetFocus
        Exit Sub
    End If
    If Val(txtTmpSize.Text) < 50 Then
        MsgBox "��ʱ�ռ��С���ô���", vbExclamation, "��ʾ"
        txtTmpSize.Text = 50
        If txtTmpSize.Visible Then txtTmpSize.SetFocus
        Exit Sub
    End If

    If Trim(lblSqlFile.Caption) = "" Then
        MsgBox "δָ�������߽ű��ļ������ܼ�����", vbExclamation, "��ʾ"
        cmdSqlFile.SetFocus
        Exit Sub
    End If

    If Trim(lblRegFile.Caption) = "" Then
        MsgBox "δִ��ע����Ȩ�ļ������ܼ�����", vbExclamation, "��ʾ"
        cmdRegFile.SetFocus
        Exit Sub
    End If
    
    If MsgBox("���ߴ������̽������ϳ���ʱ�䣬" & vbCr & "�벻Ҫ�����жϳ�������С�" & vbCr & vbCr & "������", vbQuestion + vbYesNo, "��ʾ") = vbNo Then Exit Sub
    
    Me.txtPwd.Enabled = False
    fraTbs.Enabled = False
    cmdSqlFile.Enabled = False
    cmdRegFile.Enabled = False
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    If Not CheckCBOPars Then Exit Sub
    If svrCreate(lblSqlFile.Caption) = True Then
        MsgBox "�����߳ɹ�������", vbExclamation, "��ʾ"
        If Not sysRegist(Me.lblRegFile.Caption) Then
            MsgBox "ϵͳע����Ȩ���������µ�¼����ע����Ȩ��", vbInformation, "��ʾ"
        End If
        Me.txtPwd.Enabled = True
        fraTbs.Enabled = True
        cmdSqlFile.Enabled = True
        cmdRegFile.Enabled = True
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
        Unload Me
    Else
        MsgBox "�������̷�������ϵͳ���Զ�����Ѿ�ִ�еĲ�����", vbExclamation, "��ʾ"
        Call svrRemove
        Me.txtPwd.Enabled = True
        fraTbs.Enabled = True
        cmdSqlFile.Enabled = True
        cmdRegFile.Enabled = True
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
    End If

End Sub

Private Sub Form_Load()
    
    
    imgSetup.Top = PicSetup.ScaleTop
    imgSetup.Left = PicSetup.ScaleLeft
    imgSetup.Height = PicSetup.ScaleHeight
    imgSetup.Width = PicSetup.ScaleWidth
    
    pgbState.Left = stbThis.Panels(3).Left + 90
    pgbState.Width = stbThis.Panels(4).Left - pgbState.Left - 90
    pgbState.Top = stbThis.Top + stbThis.Height / 3
    
    With rsTemp
        .Filter = 0
        If .State = adStateOpen Then .Close
        strSQL = "SELECT NAME from V$DATAFILE where ROWNUM<2 order by CREATION_TIME"
        .Open strSQL, gcnOracle, adOpenKeyset
        If .EOF Or .BOF Then
            mstrTbsPath = "C:\"
        Else
            For lngCount = Len(!name) To 2 Step -1
                If Mid(!name, lngCount, 1) = "\" Or Mid(!name, lngCount, 1) = "/" Then
                    mstrTbsPath = Left(!name, lngCount)
                    Exit For
                End If
            Next
        End If
    End With
    
    txtTbsFile.Text = mstrTbsPath & lblTbsName.Caption & ".DBF"
    txtTmpFile.Text = mstrTbsPath & lblTmpName.Caption & ".DBF"
    
    If Dir(App.Path & "\Tools\" & IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "zlRegist.plb")) <> "" And Dir(App.Path & "\Tools\zlServer.Sql") <> "" Then
        lblSqlFile.Caption = App.Path & "\Tools\zlServer.Sql"
    End If
    
    Me.fltFile.Path = App.Path
    Me.fltFile.Pattern = "*.zcr"
    If Me.fltFile.ListCount > 0 Then
        Me.lblRegFile.Caption = App.Path & "\" & Me.fltFile.List(0)
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdOK.Enabled = False Then
        Cancel = 1
        Exit Sub
    End If
    Set mclsRunScript = Nothing
    Set mobjFiles = Nothing
    Set mobjText = Nothing
End Sub

Private Sub tbsTbs_Click()
    If tbsTbs.Tabs(1).Selected Then
        lblTbsName.Visible = True
        txtTbsFile.Visible = True
        txtTbsSize.Visible = True
        
        lblTmpName.Visible = False
        txtTmpFile.Visible = False
        txtTmpSize.Visible = False
        
    ElseIf tbsTbs.Tabs(2).Selected Then
        lblTbsName.Visible = False
        txtTbsFile.Visible = False
        txtTbsSize.Visible = False
        
        lblTmpName.Visible = True
        txtTmpFile.Visible = True
        txtTmpSize.Visible = True
        
    End If

End Sub

Private Sub txtPWD_GotFocus()
    Me.txtPwd.SelStart = 0: Me.txtPwd.SelLength = 100
End Sub


'-----------------------------------------------------
'����Ϊ�ڲ�ͨ�ú���������
'-----------------------------------------------------
Private Function svrCreate(strSqlFile As String) As Boolean
    '----------------------------------
    '���ܣ����ϵͳ�İ�װ����
    '    �����������ݿռ�
    '    ��������������
    '    �����������ݶ���
    '    ��������ͬ��ʣ�����publicȨ��
    '----------------------------------
    Dim intVer As Integer
    Dim strRegFunFile As String
    Dim blnJSONRemain As Boolean
    
    '������ռ估�ع���
    stbThis.Panels(2).Text = "��������ȱʡ�ռ䡭"
    If CreateTbs(lblTbsName.Caption, txtTbsFile.Text, txtTbsSize.Text, True, False, False, 1) = 2 Then GoTo ErrHand
    
    intVer = GetOracleVersion(, True)
    If intVer >= 9 Then
        'Oracle9i�汾,�û�����ʱ�ռ�ֻ���Ǳ��ع�����ʱ��ռ䣻�Ҳ���Ҫ���������ع���
        DoEvents
        stbThis.Panels(2).Text = "����������ʱ�ռ䡭"
        If CreateTbs(lblTmpName.Caption, txtTmpFile.Text, txtTmpSize.Text, True, True, False, 1) = 2 Then GoTo ErrHand
    Else
        'Oracle8i���°汾
        DoEvents
        stbThis.Panels(2).Text = "����������ʱ�ռ䡭"
        If CreateTbs(lblTmpName.Caption, txtTmpFile.Text, txtTmpSize.Text, True, True, False, 1) = 2 Then GoTo ErrHand
    
        DoEvents
        stbThis.Panels(2).Text = "���������ع��Ρ�"
        err = 0
        On Error Resume Next
        strSQL = "create public rollback segment rbs_ZLTOOLS tablespace RBS"
        gcnOracle.Execute strSQL
        
        If err <> 0 Then
            err = 0
            On Error GoTo ErrHand
            '���ع����ռ�,������ָ���εĴ洢����
            strSQL = "create public rollback segment rbs_ZLTOOLS tablespace " & lblTbsName.Caption
            gcnOracle.Execute strSQL
        End If
        strSQL = "alter rollback segment rbs_ZLTOOLS online"
        gcnOracle.Execute strSQL
    End If
    
    '----------------------------------------------
    '��������������
    stbThis.Panels(2).Text = "�������������ߡ�"
    err = 0
    On Error Resume Next
    gcnOracle.Execute "create user ZLTOOLS identified by " & txtPwd.Text
    If err <> 0 Then
        MsgBox "�޷������������������ߣ�����" & vbNewLine & err.Description, vbExclamation, "��ʾ"
        
        gcnOracle.Execute "drop tablespace " & Trim(lblTbsName.Caption) & " including contents and datafiles cascade constraints"
        gcnOracle.Execute "drop tablespace " & Trim(lblTmpName.Caption) & " including contents and datafiles cascade constraints"
        GoTo ErrHand
    End If
    
    gcnOracle.Execute "alter user ZLTOOLS DEFAULT TABLESPACE " & Trim(lblTbsName.Caption)
    gcnOracle.Execute "alter user ZLTOOLS TEMPORARY TABLESPACE " & Trim(lblTmpName.Caption)
    gcnOracle.Execute "grant Connect,Resource,UNLIMITED TABLESPACE,Create Public Synonym,Drop Public Synonym,Alter Session,Create Session,Create Synonym,Create Table,Create View,Create Sequence,Create Database Link,Create Cluster to ZLTOOLS"
    gcnOracle.Execute "grant select on Sys.v_$session to ZLTOOLS"
    gcnOracle.Execute "grant select on Sys.gv_$session to ZLTOOLS"
    gcnOracle.Execute "grant select on Sys.dba_role_privs to ZLTOOLS"

    
    If err <> 0 Then
        MsgBox "�޷������������������ߣ��������ݿ�ϵͳ����ȷ��" & vbNewLine & err.Description, vbExclamation, "��ʾ"
        GoTo ErrHand
    End If

    '----------------------------------------------
    '�����������ݶ���
    stbThis.Panels(2).Text = "��������:"
    err = 0
    On Error GoTo ErrHand
    With cnTools
        If .State = adStateOpen Then .Close
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & Trim(gstrServer), "ZLTOOLS", txtPwd.Text
    End With
    
    Call SetSQLTrace(gstrServer, "ZLTOOLS", cnTools)

    
    Set mclsRunScript = New clsRunScript
    With mclsRunScript
        Set .Connection = cnTools: .ConnectType = 1
        Call .InitGlobalPara(Me)
        Call .InitUserList(, , txtPwd.Text)
        If IsCanInstallPLJson(gobjFSO.GetParentFolderName(strSqlFile), blnJSONRemain) Then
            Call InstallPLJSON(gcnOracle, gobjFSO.GetParentFolderName(strSqlFile), mclsRunScript, blnJSONRemain)
        End If
        On Error Resume Next
        If .OpenFile(strSqlFile) = False Then
            GoTo ErrHand
        End If
        
        pgbState.value = 0
        pgbState.Visible = True
        err = 0
        On Error GoTo ErrHand
        Do While Not mclsRunScript.EOF
            pgbState.value = Int(.ProcessValue)
            err = 0
            On Error Resume Next
            If pgbState.value > 90 Then
                Debug.Print ""
            End If
            cnTools.Execute .SQLInfo.SQL
            If err <> 0 Then
                MsgBox "�����ļ�" & strSqlFile & "�д������������ִ���жϣ�" & vbCr & .SQLInfo.SQL & vbNewLine & err.Description, vbExclamation, "��ʾ"
                GoTo ErrHand
            End If
            err = 0
            On Error GoTo ErrHand
            DoEvents
            .ReadNextSQL
        Loop
    End With
        '----------------------------------------------
        'ͨ��Shell��ʽ��������Ȩ��֤����
        
    stbThis.Panels(2).Text = "����ִ�нű���"
    strRegFunFile = Mid(strSqlFile, 1, Len(strSqlFile) - InStr(1, StrReverse(strSqlFile), "\") + 1) & IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "zlRegist.plb")
    
    If Not RunRegistFile(Me, cnTools, Trim(txtPwd.Text), gstrServer, strRegFunFile) Then
        GoTo ErrHand
    End If
    
    With rsTemp
        If .State = adStateOpen Then .Close
        strSQL = "Select 1 From User_Objects Where Object_Type = 'FUNCTION' And Object_Name = '" & UCase("f_Reg_Audit") & "' and status='VALID'"
        .Open strSQL, cnTools
        If .RecordCount = 0 Then GoTo ErrHand
    End With
    
    '----------------------------------------------
    '��������ͬ��ʣ�����publicȨ��
    stbThis.Panels(2).Text = "��Ȩ����:"
    pgbState.Visible = False
    Call ReGrantForTools(cnTools)
    cnTools.Close
    svrCreate = True
    Exit Function

ErrHand:
    If cnTools.State = adStateOpen Then cnTools.Close
    pgbState.Visible = False
    svrCreate = False
End Function

Private Function svrRemove() As Boolean
    '----------------------------------
    '���ܣ�ɾ���Ѿ��İ�װ����
    '----------------------------------
    Dim strSpaces As String, strFiles As String, aryFile() As String, strErrInfo As String
    Dim strStep As String, aryStep() As String
    Dim lngRowH As Long, intVer As Integer
    
    strFiles = ""
    With rsTemp
        .Filter = 0
        If .State = adStateOpen Then .Close
        strSQL = "select F.NAME " & _
                " from V$TABLESPACE T,V$DATAFILE F " & _
                " where T.TS#=F.TS# " & _
                "       and T.NAME in('" & UCase(lblTbsName.Caption) & "','" & UCase(lblTbsName.Caption) & "')"
        .Open strSQL, gcnOracle
        Do While Not .EOF
            strFiles = strFiles & ";" & .Fields(0).value
            DoEvents
            .MoveNext
        Loop
    End With
    
    err = 0
    On Error Resume Next
    stbThis.Panels(2).Text = "ɾ�������ߡ�"
    Do
        gcnOracle.Execute "drop user ZLTOOLS cascade"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open "select * from all_users where username='ZLTOOLS'", gcnOracle
        If rsTemp.EOF Then Exit Do
        lngCount = lngCount + 1
        DoEvents
        If lngCount > 10000 Then
            strErrInfo = strErrInfo & vbCr & "�û�:ZLTOOLS"
            Exit Do
        End If
    Loop
    
    stbThis.Panels(2).Text = "ɾ�����߿ռ估���ļ���"
    intVer = GetOracleVersion(, True)
    If intVer < 9 Then
        gcnOracle.Execute "alter rollback segment rbs_ZLTOOLS offline"
        gcnOracle.Execute "drop rollback segment rbs_ZLTOOLS"
    End If
    
    gcnOracle.Execute "alter tablespace " & lblTmpName.Caption & " offline"
    gcnOracle.Execute "alter tablespace " & lblTbsName.Caption & " offline"
    
    gcnOracle.Execute "drop tablespace " & lblTmpName.Caption & " including contents and datafiles cascade constraints"
    gcnOracle.Execute "drop tablespace " & lblTbsName.Caption & " including contents and datafiles cascade constraints"
    
    
    aryFile = Split(Mid(strFiles, 2), ";")
    For lngCount = 0 To UBound(aryFile)
        err = 0
        mobjFiles.DeleteFile aryFile(lngCount), True
        If err <> 0 Then
            strErrInfo = strErrInfo & vbCr & "�ļ���" & aryFile(lngCount)
        End If
    Next
    
    stbThis.Panels(2).Text = ""
    If strErrInfo <> "" Then
        MsgBox "�����������ݿ��,�ֹ�ɾ���������ݣ�" & strErrInfo, vbExclamation, "��ʾ"
    Else
        MsgBox "����Ӳ�̿ռ�����ݿ�ϵͳ��ȷ����������²���", vbExclamation, "��ʾ"
    End If
End Function

Private Function CreateTbs(TbsName As String, TbsFile As String, TbsSize As Integer, Optional AutoExtend As Boolean, _
     Optional Temp As Boolean, Optional AutoAllocate As Boolean, Optional ExtentSize As Integer) As Byte
    '----------------------------------------------
    '���ܣ�ϵͳ�û�,���ݲ���������ռ�,�̶�Ϊ���ع�������(8i��ǰ��֧��,��ʱֻ�ܴ����ֵ��������)
    '       ������漰LOB�ֶε�ԭ��,������ASSM��ռ�(��9i����֧��,SEGMENT SPACE MANAGEMENT AUTO)
    '������
    '   TbsName:��ռ�����
    '   TbsFile:��ռ��ļ�
    '   TbsSize:��ռ��С(MΪ��λ)
    '   Extend:�Ƿ��Զ�������,����ͳһ��Χ�ߴ�
    '   ExtentSize:ͳһ���ߴ�,��ʱ��ռ����ָ���ߴ�(OracleȱʡΪ1M)
    '   Temp:�Ƿ�Ϊ��ʱ��ռ�
    '���أ�1-�����ɹ���2-��ռ��Ѿ����ڣ�3-����ʧ��
    '----------------------------------------------
    DoEvents
    If Temp Then
        gstrSQL = "CREATE TEMPORARY TABLESPACE " & TbsName & " TEMPFILE '" & TbsFile & "'"
    Else
        gstrSQL = "CREATE TABLESPACE " & TbsName & " DATAFILE '" & TbsFile & "'"
    End If
    gstrSQL = gstrSQL & _
            " SIZE " & TbsSize & "M REUSE " & _
             IIf(AutoExtend, "AUTOEXTEND ON NEXT " & IIf(TbsSize \ 10 = 0, 1, TbsSize \ 10) & "M", "") & _
            " EXTENT MANAGEMENT LOCAL " & _
                IIf(AutoAllocate And Not Temp, " AUTOALLOCATE", " UNIFORM SIZE " & IIf(ExtentSize = 0, "1", ExtentSize) & "M")
    
    err = 0
    On Error Resume Next
    gcnOracle.Execute gstrSQL
    DoEvents
    If err = 0 Then
        CreateTbs = 1
    ElseIf gcnOracle.Errors.Count > 0 Then
        MsgBox gcnOracle.Errors(0).Description & _
            IIf(InStr(1, gcnOracle.Errors(0).Description, "00406") > 0, vbCrLf & "���޸�Oracle�������ļ���Compatible����Ϊ8.1.5����", ""), _
            vbExclamation, "��ʾ"
        CreateTbs = 2
    Else
        MsgBox "��ռ�" & TbsName & "�޷�������������̴�С��", vbExclamation, "��ʾ"
        CreateTbs = 2
    End If

End Function

Private Function sysRegist(strRegFile As String) As Boolean
    '----------------------------------
    '���ܣ�ϵͳע��
    '----------------------------------
    stbThis.Panels(2).Text = "ϵͳע����Ȩ��"
    
    'д����ʱ������֤
    err = 0: On Error GoTo ErrHand
    Me.MousePointer = vbHourglass
    
    If gobjRegister.zlRegBuild(strRegFile, pgbState) = False Then GoTo ErrHand
    
    Me.MousePointer = vbDefault
    
    If gobjRegister.zlRegCheck(True) <> "" Then GoTo ErrHand
    
    '��ʽд��
    gcnOracle.Execute "call zltools.p_Reg_Apply()", , adCmdText
    
    sysRegist = True
    Exit Function
ErrHand:
    Me.MousePointer = vbDefault
    sysRegist = False
End Function

