VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmSurety 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "סԺ������Ϣ����"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmSurety.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7575
   Begin VB.PictureBox PicDeposit 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   150
      ScaleHeight     =   3090
      ScaleWidth      =   5790
      TabIndex        =   29
      Top             =   3510
      Width           =   5790
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
         Height          =   2265
         Left            =   0
         TabIndex        =   31
         Top             =   330
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   3995
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorBkg    =   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblDeposit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ���ܶ"
         Height          =   180
         Left            =   45
         TabIndex        =   30
         Top             =   75
         Width           =   900
      End
   End
   Begin VB.Frame fraPati 
      Height          =   960
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   7350
      Begin VB.TextBox txtPatient 
         Height          =   300
         Left            =   1350
         TabIndex        =   3
         Top             =   225
         Width           =   1275
      End
      Begin VB.CommandButton cmdPati 
         Height          =   300
         Left            =   2625
         Picture         =   "frmSurety.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����(F2)"
         Top             =   225
         Width           =   300
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   300
         Left            =   720
         TabIndex        =   2
         ToolTipText     =   "��ݼ�F4"
         Top             =   225
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   529
         Appearance      =   2
         IDKindStr       =   $"frmSurety.frx":0914
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lblCur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ƹ��ʽ��"
         Height          =   180
         Left            =   5085
         TabIndex        =   33
         Top             =   285
         Width           =   1260
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�ȼ���"
         Height          =   180
         Left            =   5085
         TabIndex        =   32
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2985
         TabIndex        =   5
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Left            =   3960
         TabIndex        =   6
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ�ţ�"
         Height          =   180
         Left            =   330
         TabIndex        =   7
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ң�"
         Height          =   180
         Left            =   2325
         TabIndex        =   8
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lblBed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ţ�"
         Height          =   180
         Left            =   3960
         TabIndex        =   9
         Top             =   630
         Width           =   540
      End
   End
   Begin VB.PictureBox picSurety 
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   120
      ScaleHeight     =   3900
      ScaleWidth      =   7425
      TabIndex        =   27
      Top             =   1170
      Width           =   7425
      Begin VB.Frame fraEdit 
         Caption         =   "��Ϣ����"
         Height          =   1095
         Left            =   0
         TabIndex        =   10
         Top             =   15
         Width           =   7335
         Begin VB.TextBox txtWarrantM 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2760
            MaxLength       =   9
            TabIndex        =   14
            Top             =   360
            Width           =   1005
         End
         Begin VB.TextBox txtWarrantP 
            Height          =   300
            Left            =   840
            MaxLength       =   100
            TabIndex        =   12
            Top             =   360
            Width           =   1005
         End
         Begin VB.CheckBox chkUnlimit 
            Caption         =   "���޶��"
            Height          =   255
            Left            =   2760
            TabIndex        =   18
            ToolTipText     =   "���޵�����ʱ�������õ���ʱ��"
            Top             =   720
            Width           =   1050
         End
         Begin VB.CheckBox chkWarrantL 
            Caption         =   "��ʱ����"
            Height          =   255
            Left            =   840
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   720
            Width           =   1050
         End
         Begin VB.TextBox txtReason 
            Height          =   300
            Left            =   5040
            MaxLength       =   50
            TabIndex        =   20
            Top             =   720
            Width           =   2010
         End
         Begin MSComCtl2.DTPicker dtpWarrantT 
            Height          =   300
            Left            =   5040
            TabIndex        =   16
            Top             =   345
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   93323267
            CurrentDate     =   38915.6041666667
         End
         Begin VB.Label lblWarrantM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   2160
            TabIndex        =   13
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lblWarrantP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   240
            TabIndex        =   11
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lblWarrantT 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   4140
            TabIndex        =   15
            ToolTipText     =   "��Ժ���˲���ʹ��ʱ�޵���"
            Top             =   450
            Width           =   720
         End
         Begin VB.Label lblReason 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ԭ��"
            Height          =   180
            Left            =   4140
            TabIndex        =   19
            Top             =   780
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdAdd 
         Cancel          =   -1  'True
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "�������һ��������¼���ڻ�û����������ʱ����������"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "�޸�(&M)"
         Height          =   350
         Left            =   1350
         TabIndex        =   22
         ToolTipText     =   "ֻ�����޸����һ��������¼"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   2450
         TabIndex        =   23
         ToolTipText     =   "ֻ����ɾ�����һ��������¼"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�(&X)"
         Height          =   350
         Left            =   6000
         TabIndex        =   24
         ToolTipText     =   "(F9)�˳�"
         Top             =   1200
         Width           =   1100
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
         Height          =   2265
         Left            =   0
         TabIndex        =   25
         Top             =   1680
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   3995
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483645
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorBkg    =   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   5145
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9499
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3775
            MinWidth        =   3775
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
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   3945
      Left            =   555
      TabIndex        =   28
      Top             =   1035
      Width           =   3795
      _Version        =   589884
      _ExtentX        =   6694
      _ExtentY        =   6959
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmSurety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mlng����ID As Long
Public mbln��Ժ���� As Boolean
Public mstrPrivs As String
Private mlng��ҳID As Long      '��Ժ����Ϊ��ǰסԺ�Ǽǵ���ҳID

Private mrsInfo As New ADODB.Recordset
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mobjSquareCard As Object
Private mblnDefaultPassInputCardNo As Boolean 'ȱʡˢ���Ƿ��������뿨��
Private mblnNotClick As Boolean
Private mblnFirst As Boolean
Private mstr�������� As String

Private Sub chkUnlimit_Click()
     '���޵����������ʱ����
    If chkUnlimit.Value = 1 And IsNull(dtpWarrantT.Value) Then
        dtpWarrantT.Value = DateAdd("d", 3, dtpWarrantT.MinDate)
    End If
    chkWarrantL.Enabled = Not (chkUnlimit.Value = 1)
    txtWarrantM.Enabled = Not (chkUnlimit.Value = 1)
    
    If chkUnlimit.Value = 1 Then
        txtWarrantM.Text = "999999999":  txtWarrantM.BackColor = vbInactiveCaptionText
    Else
        txtWarrantM.Text = "": txtWarrantM.BackColor = vbWhite
    End If
End Sub

Private Sub chkWarrantL_Click()
    If chkWarrantL.Value = 1 Then
        dtpWarrantT.CheckBox = True: dtpWarrantT.CustomFormat = "yyyy-MM-dd HH:mm"
        dtpWarrantT.Value = Null
        chkUnlimit.Value = 0        'ֵ�ı�ʱ����ʽ����click�¼�
    End If
    chkUnlimit.Enabled = Not (chkWarrantL.Value = 1) And mbln��Ժ����
    dtpWarrantT.Enabled = Not (chkWarrantL.Value = 1) And mbln��Ժ����
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    Dim str�Ǽ�ʱ�� As String
    Dim strɾ����־ As String
    Dim blnOk As Boolean
    
    blnOk = True
    If mrsInfo Is Nothing Then
        blnOk = False
    ElseIf mrsInfo.State = adStateClosed Then
        blnOk = False
    End If
    
    If blnOk = False Then
        stbThis.Panels(1).Text = "û��ȷ��Ҫ���е����Ĳ���!"
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    End If
    
    '����21368 by lesfeng 2010-08-02
    strɾ����־ = Trim(msh.TextMatrix(msh.Row, GetColNum("ɾ����־")))
    If strɾ����־ = "ɾ��" Then
        MsgBox "����������¼�Ѿ�Ϊɾ����ǣ����ܽ���ɾ����ǲ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("ȷʵҪ���б�Ǵ���������¼Ϊɾ�������?" & vbCrLf & vbCrLf & "ע��,ɾ����Ǻ󣬵�ǰ�������᲻�ָܻ�!" _
        , vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errH
    
    If Trim(msh.TextMatrix(msh.Row, GetColNum("�Ǽ�ʱ��"))) = "" Then
        str�Ǽ�ʱ�� = "NULL"
    Else
        str�Ǽ�ʱ�� = To_Date(Trim(msh.TextMatrix(msh.Row, GetColNum("�Ǽ�ʱ��"))))
    End If
    '����21368 by lesfeng 2010-08-02
    strSQL = "zl_���˵�����¼_delete(" & mlng����ID & "," & mlng��ҳID & ",NULL," & str�Ǽ�ʱ�� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    stbThis.Panels(1).Text = "ɾ�������ɹ�!"
    Call LoadSurety
    
    If cmdExit.Enabled Then cmdExit.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdModify_Click()
    Dim strSQL As String, str������ As String, str����ʱ�� As String
    Dim str�Ǽ�ʱ�� As String
    Dim strɾ����־ As String
    Dim blnOk As Boolean
    'ֻ���޸ĵ�ǰѡ�в�����Ч�ĵ�����¼
    
    
    If cmdModify.Caption = "�޸�(&M)" Then
        
        blnOk = True
        If mrsInfo Is Nothing Then
            blnOk = False
        ElseIf mrsInfo.State = adStateClosed Then
            blnOk = False
        End If
        
        If blnOk = False Then
            stbThis.Panels(1).Text = "û��ȷ��Ҫ���е����Ĳ���!"
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Exit Sub
        End If
    
    '��ȡ�޸���Ϣ
        If msh.TextMatrix(msh.Row, GetColNum("������")) = "" Then
            stbThis.Panels(1).Text = "û�п����޸ĵĵ�����Ϣ!"
            Exit Sub
        End If
        '����21368 by lesfeng 2010-08-02
        strɾ����־ = Trim(msh.TextMatrix(msh.Row, GetColNum("ɾ����־")))
        If strɾ����־ = "ɾ��" Then
            MsgBox "����������¼�Ѿ�Ϊɾ����ǣ����ܽ����޸Ĳ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        cmdModify.Caption = "����(&S)"
        cmdAdd.Enabled = False
        cmdDel.Enabled = False
        cmdExit.Caption = "ȡ��(&C)"
        fraEdit.Enabled = True
        
        With msh
            txtWarrantP.Text = Trim(.TextMatrix(.Row, GetColNum("������")))
            If .TextMatrix(.Row, GetColNum("������")) = "����" Then
                chkUnlimit.Value = 1    'ֵ��ͬʱ��ʽ����click�¼�
                txtWarrantM.Text = "999999999"
            Else
                chkUnlimit.Value = 0
                txtWarrantM.Text = Val(.TextMatrix(.Row, GetColNum("������")))
            End If
            
            If IsDate(.TextMatrix(.Row, GetColNum("����ʱ��"))) Then
                dtpWarrantT.CheckBox = True: dtpWarrantT.CustomFormat = "yyyy-MM-dd HH:mm"
                dtpWarrantT.Value = CDate(.TextMatrix(.Row, GetColNum("����ʱ��")))
            Else
                dtpWarrantT.CheckBox = True: dtpWarrantT.CustomFormat = "yyyy-MM-dd HH:mm" '������ɼ��������ִ�л����
                dtpWarrantT.Value = Null
            End If
            
            chkWarrantL.Value = IIf(.TextMatrix(.Row, GetColNum("��ʱ����")) = "��", 1, 0)
            If txtWarrantP.Enabled Then txtWarrantP.SetFocus
            txtWarrantP.Tag = Trim(.TextMatrix(msh.Row, GetColNum("�Ǽ�ʱ��")))
        End With
    Else
    '�����޸Ľ��
        '1.���ݼ��
        If Not Check������Ϣ Then Exit Sub
        
        
        '�Ȼָ����水ť״̬
        cmdModify.Caption = "�޸�(&M)"
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExit.Caption = "�˳�(&X)"
        fraEdit.Enabled = True      'SetCanEdit���ٴ�����
        
        str������ = Replace(Trim(txtWarrantP.Text), "'", "''")
        str����ʱ�� = "null"
        If Not IsNull(dtpWarrantT.Value) Then str����ʱ�� = To_Date(dtpWarrantT.Value)
        str�Ǽ�ʱ�� = To_Date(txtWarrantP.Tag)
        
        '���ȼ��
        If Not CheckLen(txtWarrantP, 64) Then Exit Sub
        
        '2.���ݱ���
        On Error GoTo errH
        strSQL = "zl_���˵�����¼_update(" & mlng����ID & "," & mlng��ҳID & ",'" & str������ & "'," & _
            Val(txtWarrantM.Text) & "," & chkWarrantL.Value & ",'" & Trim(txtReason.Text) & "',NULL," & str����ʱ�� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & str�Ǽ�ʱ�� & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                
        '3.����ˢ��
        stbThis.Panels(1).Text = "�޸Ľ���ѱ���!"
        Call LoadSurety
        Call Init������Ϣ
        If cmdExit.Enabled Then cmdExit.SetFocus
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Init������Ϣ()
    Dim Datsys As Date

    txtWarrantP.Text = ""
    chkUnlimit.Enabled = mbln��Ժ����
    chkUnlimit.Value = 0            '���ֵ�б仯,����ʽ����click�¼�
    txtWarrantM.Text = ""
    txtReason.Text = ""
    
    dtpWarrantT.Enabled = mbln��Ժ����
    dtpWarrantT.CheckBox = True: dtpWarrantT.CustomFormat = "yyyy-MM-dd HH:mm" '����checkbox�ɼ���
    If dtpWarrantT.Enabled Then
        Datsys = zlDatabase.Currentdate
        dtpWarrantT.MinDate = Datsys
        dtpWarrantT.Value = DateAdd("d", 3, Datsys)
    End If
    dtpWarrantT.Value = Null
    
    chkWarrantL.Enabled = True
    chkWarrantL.Value = 0
    chkUnlimit.TabStop = True
End Sub

Public Sub InitFace()
    lblSex.Caption = "�Ա�": lblNO.Caption = "סԺ�ţ�": lblBed.Caption = "���ţ�"
    lblAge.Caption = "���䣺": lblDept.Caption = "���ң�": lblDeposit.Caption = "Ԥ���ܶ"
    lblType.Caption = "�ѱ�ȼ���": lblCur.Caption = "ҽ�Ƹ��ʽ��"
End Sub

Private Sub cmdPati_Click()
    If frmPatiSelect.ShowMe(Me) = True Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub dtpWarrantT_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    ElseIf KeyAscii = vbKeySpace Then
        If dtpWarrantT.CheckBox Then
            KeyAscii = 0
            If IsNull(dtpWarrantT.Value) Then
                dtpWarrantT.Value = DateAdd("d", 3, zlDatabase.Currentdate)
            Else
                dtpWarrantT.Value = Null
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = True Then Exit Sub
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    mblnFirst = True
End Sub

Private Sub Form_Load()
        
    Dim strSQL  As String
    Dim rsTmp As New ADODB.Recordset
    
    mblnFirst = False
    Call RestoreWinState(Me, App.ProductName)
    Call InitTabPage
    Call zlCardSquareObject
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, "", txtPatient)
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Set mobjICCard.gcnOracle = gcnOracle
    IDKind.Enabled = True
    
    If Not mobjSquareCard Is Nothing Then
        IDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    Call ClearWinInfor(True)
    
    fraEdit.Enabled = True
    If InStr(mstrPrivs, "����Ǽ�") <= 0 And InStr(mstrPrivs, "����ԤԼ") = 0 And InStr(mstrPrivs, "���ղ��˵Ǽ�") <= 0 Then
        fraEdit.Enabled = False
        cmdAdd.Visible = False
        cmdModify.Visible = False
        cmdDel.Visible = False
        Me.Caption = "סԺ������Ϣ�鿴(��ǰ�û���" & UserInfo.���� & ")"
    End If
    
    txtWarrantP.Enabled = fraEdit.Enabled
    txtWarrantP.BackColor = IIf(fraEdit.Enabled, &H80000005, &H8000000F)
    txtWarrantM.Enabled = fraEdit.Enabled
    txtWarrantM.BackColor = IIf(fraEdit.Enabled, &H80000005, &H8000000F)
    chkWarrantL.Enabled = fraEdit.Enabled
    chkUnlimit.Enabled = fraEdit.Enabled
    txtReason.Enabled = fraEdit.Enabled
    txtReason.BackColor = IIf(fraEdit.Enabled, &H80000005, &H8000000F)
    If mlng����ID > 0 Then
        txtPatient.Text = "-" & mlng����ID
        Call txtPatient_KeyPress(vbKeyReturn)
    Else
        cmdAdd.Enabled = False
    End If
End Sub

Private Sub ClearWinInfor(Optional ByVal blnClear As Boolean = False)
    Call InitFace
    Call LoadSurety(blnClear)
    Call LoadPrepay(blnClear)
    Call Init������Ϣ
End Sub

Private Sub InitTabPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ҳ�ؼ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
        
    Set objItem = tbcPage.InsertItem(1, "������Ϣ", picSurety.hWnd, 0)
    objItem.Tag = 1
    
    Set objItem = tbcPage.InsertItem(2, "Ԥ����Ϣ", PicDeposit.hWnd, 0)
    objItem.Tag = 2
    
    With tbcPage
        .Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
    GetColNum = -1
End Function

Private Function GetColNumList(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNumList = i: Exit Function
    Next
    GetColNumList = -1
End Function

Private Sub SetSuretyHeader()
    Dim strHead As String, i As Long
    strHead = ",4,300|���,4,1000|������,4,800|������,7,1250|��ʱ����,4,850|����ԭ��,4,1800|�Ǽ�ʱ��,1,1800|����ʱ��,1,1800|ɾ����־,4,850|����Ա����,4,1050|����Ա���,4,1050|ɾ������Ա����,4,1050|ɾ������Ա���,4,1050|ɾ��ʱ��,1,1800"
    With msh
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .colAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(msh, App.ProductName & "\" & Me.Name)
        
        .ForeColor = &H80000003
        .RowHeight(0) = 320
        .Redraw = True
    End With
End Sub

Private Sub SetDepositHeader()
    Dim strHead As String, i As Long
    strHead = ",4,300|����,4,1350|���ݺ�,4,1110|����,1,1200|���,1,0|�ɿ���,7,1600|����,4,1000|�տ���,1,1000"
    With mshList
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .colAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(msh, App.ProductName & "\" & Me.Name)
        
        .ForeColor = &H80000003
        .RowHeight(0) = 320
        .Redraw = True
    End With
End Sub

Private Sub GetSuretyBalance()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = _
        " Select To_char(������,'99999999990.00') as ������,Decode(��ǰ����ID,null,0,��ҳID) as ��ҳID" & _
        " From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If rsTmp.RecordCount > 0 Then
        stbThis.Panels(2).Text = "��Ч������:" & IIf(IsNull(rsTmp!������), "��", Val(Trim("" & rsTmp!������)))
        'mlng��ҳID = Val("" & rsTmp!��ҳID)
    Else
        stbThis.Panels(2).Text = ""
        'mlng��ҳID = 0
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadSurety(Optional ByVal blnClear As Boolean = False)
    Dim rsTmp As ADODB.Recordset, Curdate As Date
    Dim strSQL As String, i As Integer, lngRow As Integer, RowPageid As Integer
    Dim strɾ����־ As String
    Dim lng����ID As Long, lng��ҳID As Long
    
    On Error GoTo errH
    If mrsInfo Is Nothing Then
        lng����ID = mlng����ID
        lng��ҳID = mlng��ҳID
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = mlng����ID
        lng��ҳID = mlng��ҳID
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
        lng��ҳID = Val(Nvl(mrsInfo!��ҳID))
    End If
    stbThis.Panels(2).Text = ""
    If blnClear = True Then
        msh.Clear
        msh.Rows = 2
        msh.RowData(1) = 0
        Call SetSuretyHeader
    Else
        Curdate = zlDatabase.Currentdate
        '����21368 by lesfeng 2010-08-02
        'ɾ����־,4,850|����Ա����,4,1050|����Ա���,4,1050|ɾ������Ա����,4,1050|ɾ������Ա���,4,1050|ɾ��ʱ��,1,1800"
        strSQL = _
            "SELECT '',Decode(��ҳid, NULL, '����', '��' || ��ҳid || '��סԺ') ���, ������," & vbNewLine & _
            "       Decode(������, 999999999, '����', To_Char(������, '999999990.00')) AS ������," & vbNewLine & _
            "       Decode(��������, 1, '��', ' ') AS ��ʱ����, ����ԭ��, To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') �Ǽ�ʱ��," & vbNewLine & _
            "       To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��,decode(ɾ����־,1,'',-1,'ɾ��','') as ɾ����־," & vbNewLine & _
            "       ����Ա����,����Ա���,ɾ������Ա����,ɾ������Ա���,ɾ��ʱ��" & vbNewLine & _
            "FROM ���˵�����¼" & vbNewLine & _
            "WHERE ����id = [1] And ��ҳID=[2]" & vbNewLine & _
            "ORDER BY �Ǽ�ʱ�� DESC"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        If rsTmp.RecordCount > 0 Then
            Set msh.DataSource = rsTmp
            Do While Not rsTmp.EOF
                msh.RowData(rsTmp.AbsolutePosition) = lng����ID
            rsTmp.MoveNext
            Loop
        Else
            msh.Clear
            msh.Rows = 2
        End If
        Call SetSuretyHeader
        Call GetSuretyBalance
        For lngRow = 1 To msh.Rows - 1
            If UBound(Split(Trim(msh.TextMatrix(lngRow, GetColNum("���"))), "��סԺ")) > 0 Then 'ȡ��ѡ������ҳID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(lngRow, GetColNum("���"))), "��סԺ")(0), "��")(1))
            Else
                RowPageid = 0
            End If
            '����21368 by lesfeng 2010-08-02
            strɾ����־ = Trim(msh.TextMatrix(lngRow, GetColNum("ɾ����־")))
            
            If lng��ҳID = RowPageid And (Trim(msh.TextMatrix(lngRow, GetColNum("����ʱ��"))) = "" Or Trim(msh.TextMatrix(lngRow, GetColNum("����ʱ��"))) > Curdate) Then
                msh.Row = lngRow
                For i = 0 To msh.Cols - 1
                    msh.Col = i
                    '����21368 by lesfeng 2010-08-02
                    If strɾ����־ = "" Then
                        msh.CellForeColor = &HC00000
                    Else
                        msh.CellForeColor = &HFF&
                    End If
                Next
            Else
                 For i = 0 To msh.Cols - 1
                    msh.Col = i
                    '����21368 by lesfeng 2010-08-02
                    If strɾ����־ = "" Then
                    Else
                        msh.CellForeColor = &HFF&
                    End If
                Next
            End If
            
        Next lngRow
    End If
    msh.Row = 1
    msh.Col = 0: msh.ColSel = msh.Cols - 1
    Call msh_EnterCell
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadPrepay(Optional ByVal blnClear As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ʷ��Ԥ������
    '����:������
    '����:2013-03-11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lngRow As Long
    Dim rsMoney As ADODB.Recordset
    Dim lng����ID As Long, lng��ҳID As Long
    
    If mrsInfo Is Nothing Then
        lng����ID = mlng����ID
        lng��ҳID = mlng��ҳID
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = mlng����ID
        lng��ҳID = mlng��ҳID
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
        lng��ҳID = Val(Nvl(mrsInfo!��ҳID))
    End If
    
    On Error GoTo errHandle
    
    If blnClear = True Then
        mshList.Clear
        mshList.Rows = 2
        Call SetDepositHeader
    Else
        '������ʷ�ɿ���ϸ�嵥
        strSQL = _
        " Select '',Ltrim(To_Char(A.�տ�ʱ��,'YYYY-MM-DD')) as ����,A.NO as ���ݺ�,B.���� as ����,A.���, " & _
        " Ltrim(To_Char(A.���,'9,999,999,990.00')) as �ɿ���,A.���㷽ʽ as ����,A.����Ա���� as �տ��� " & _
        " From ����Ԥ����¼ A,���ű� B" & _
        " Where A.����ID=B.ID(+) And A.��¼����=1 And A.����ID=[1] And A.��ҳID=[2] And A.Ԥ�����=[3] " & _
        " Order by A.�տ�ʱ�� Desc"
        
        Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, 2)
        If rsMoney.RecordCount > 0 Then
            Set mshList.DataSource = rsMoney
        Else
            mshList.Clear
            mshList.Rows = 2
        End If
        Call SetDepositHeader
    End If
    If mshList.Rows > 1 Then
        mshList.Row = 1: mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Check������Ϣ() As Boolean
    Check������Ϣ = True
    
    If mrsInfo Is Nothing Then
        Check������Ϣ = False
    ElseIf mrsInfo.State = adStateClosed Then
        Check������Ϣ = False
    End If
    
    If Check������Ϣ = False Then
        stbThis.Panels(1).Text = "û��ȷ��Ҫ���е����Ĳ���!"
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    If Trim(txtWarrantP.Text) = "" Then
        stbThis.Panels(1).Text = "�����뵣��������,�����˲���Ϊ��!"
        If txtWarrantP.Enabled Then txtWarrantP.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
    If Not IsNumeric(txtWarrantM.Text) Then
        stbThis.Panels(1).Text = "��������ȷ�ĵ�����,������Ҫ������ֵ!"
        If txtWarrantM.Enabled Then txtWarrantM.SetFocus
        Check������Ϣ = False
        Exit Function
    ElseIf Val(txtWarrantM.Text) = 0 Then
        stbThis.Panels(1).Text = "�����뵣����,�������Ϊ��!"
        If txtWarrantM.Enabled Then txtWarrantM.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
    If chkWarrantL.Value = 1 Then
        If Not IsNull(dtpWarrantT.Value) Or chkUnlimit.Value = 1 Then
            stbThis.Panels(1).Text = "��ʱ�������������õ���ʱ�޻��޵�����!"
            If chkWarrantL.Enabled Then chkWarrantL.SetFocus
            Check������Ϣ = False
            Exit Function
        End If
    End If
    
    If zlCommFun.ActualLen(Trim(txtReason.Text)) > 50 Then
        stbThis.Panels(1).Text = "����ԭ�������������� 25 �����ֻ� 50 ���ַ���"
        txtReason.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
End Function

Private Sub cmdAdd_Click()
    Dim str������ As String, str����ʱ�� As String
    Dim strSQL As String, i As Integer, Curdate As Date, blnδ���� As Boolean, bln��ʱ As Boolean, RowPageid As Integer
    Dim strɾ����־ As String
    
    '1.���ݼ��
    If Not Check������Ϣ Then Exit Sub
    
    Curdate = zlDatabase.Currentdate
    
    For i = 1 To msh.Rows - 1 '�жϱ���סԺδ���ڵĵ�����¼��������ʾ
         If Trim(msh.TextMatrix(i, GetColNum("���"))) <> "" Then
            If UBound(Split(Trim(msh.TextMatrix(i, GetColNum("���"))), "��סԺ")) > 0 Then 'ȡ��ѡ������ҳID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(i, GetColNum("���"))), "��סԺ")(0), "��")(1))
            Else
                RowPageid = 0
            End If
            If mlng��ҳID = RowPageid Then
                '����21368 by lesfeng 2010-08-02
                strɾ����־ = Trim(msh.TextMatrix(i, GetColNum("ɾ����־")))
               If (Trim(Nvl(msh.TextMatrix(i, GetColNum("����ʱ��")))) = "" Or Nvl(msh.TextMatrix(i, GetColNum("����ʱ��"))) > Curdate) And strɾ����־ = "" Then
                   bln��ʱ = Nvl(msh.TextMatrix(i, GetColNum("��ʱ����"))) = "��"
                   blnδ���� = True: Exit For
               End If
            End If
        End If
    Next
    
    If blnδ���� Then
        If MsgBox("����δ���ڵ�" & IIf(bln��ʱ, "��ʱ", "") & "������¼����������" & IIf(bln��ʱ, "��֮ǰ����ʱ�����Զ�ʧЧ", "�ۼƵ���") & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
        
    str������ = Replace(Trim(txtWarrantP.Text), "'", "''")
    str����ʱ�� = "null"
    If Not IsNull(dtpWarrantT.Value) Then str����ʱ�� = "To_Date('" & Format(dtpWarrantT.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    '���ȼ��
    If Not CheckLen(txtWarrantP, 64) Then Exit Sub
    
    '2.���ݱ���
    On Error GoTo errH
    
    strSQL = "zl_���˵�����¼_insert(" & mlng����ID & "," & mlng��ҳID & ",'" & str������ & "'," & _
        Val(txtWarrantM.Text) & "," & chkWarrantL.Value & ",'" & Trim(txtReason.Text) & "',Null," & str����ʱ�� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '3.����ˢ��
    stbThis.Panels(1).Text = "������Ϣ�ѱ���!"
    Call LoadSurety
    Call Init������Ϣ
    
    If cmdExit.Enabled Then cmdExit.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    
    If cmdExit.Caption = "ȡ��(&C)" Then
        cmdModify.Caption = "�޸�(&M)"
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExit.Caption = "�˳�(&X)"
        fraEdit.Enabled = True      'SetCanEdit���ٴ�����
       
        'ˢ������,���ǲ�������
        stbThis.Panels(1).Text = ""
        Call LoadSurety
        Call Init������Ϣ
    Else
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim obj As Control
    Select Case KeyCode
    Case vbKeyEscape
        Call cmdExit_Click
    Case vbKeyF2
        Call cmdPati_Click
    Case vbKeyF4
        If Shift = vbCtrlMask And IDKind.Enabled Then
            Dim intIndex As Integer
            intIndex = IDKind.GetKindIndex("IC����")
            If intIndex <= 0 Then Exit Sub
             IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
        End If
    Case vbKeyF11
        If txtPatient.Enabled And Not txtPatient.Locked Then txtPatient.SetFocus
    Case vbKeyReturn
        Set obj = Me.ActiveControl
        If InStr(1, ",txtWarrantP,txtWarrantM,dtpWarrantT,chkWarrantL,chkUnlimit,txtReason,", "," & obj.Name & ",") > 0 Then
           ' Call zlCommFun.PressKey(vbKeyTab)
        End If
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With tbcPage
        .Left = fraPati.Left
        .Top = fraPati.Top + fraPati.Height
        .width = fraPati.width
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
    
    PicDeposit.width = picSurety.width
    PicDeposit.Height = picSurety.Height
    
    With msh
        .width = picSurety.ScaleWidth
        .Height = picSurety.ScaleHeight - .Top
    End With
    
    With lblDeposit
        .Left = 60
        .Top = 60
    End With
    
    With mshList
        .Top = lblDeposit.Top + lblDeposit.Height + 60
        .Left = 0
        .width = msh.width
        .Height = PicDeposit.ScaleHeight - .Top
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdModify.Caption = "����(&S)" Then
        If MsgBox("��ǰ�޸ĵ���Ϣδ����,ȷʵҪ�˳���?", vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then Cancel = 1
    End If
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    Call zlCardSquareObject(True)
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
               Set mobjICCard = New clsICCard
               Call mobjICCard.SetParent(Me.hWnd)
               Set mobjICCard.gcnOracle = gcnOracle
        End If
           If Not mobjICCard Is Nothing Then
               txtPatient.Text = mobjICCard.Read_Card()
               If txtPatient.Text <> "" Then
                   Call txtPatient_KeyPress(vbKeyReturn)
               End If
           End If
           Exit Sub
    End If
     
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If mobjSquareCard.zlReadCard(Me, glngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub
 
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    Call txtPatient_GotFocus
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    Dim lngPreIDKind As Long, lngIndex As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        lngIndex = IDKind.GetKindIndex("IC����")
        If lngIndex >= 0 Then IDKind.IDKind = lngIndex
        txtPatient.Text = strCardNO
        Call txtPatient_KeyPress(vbKeyReturn)
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long, lngIndex As Long
    
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        lngIndex = IDKind.GetKindIndex("���֤��")
        If lngIndex >= 0 Then IDKind.IDKind = lngIndex
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub

Private Sub msh_EnterCell()
    Dim str����ʱ�� As String
    Dim Datsys As Date, RowPageid As Integer
    Dim strɾ����־ As String
    
    If Val(msh.RowData(msh.Row)) <= 0 Then
        stbThis.Panels(1).Text = ""
        cmdModify.Enabled = False
        cmdDel.Enabled = False
        Exit Sub
    End If
   '��ǰ����ҳ�벡����ҳ��ͬʱ�������޸�ɾ��,�ѹ��ڲ������޸�ɾ��
    Datsys = zlDatabase.Currentdate
    
    '����21368 by lesfeng 2010-08-02
    strɾ����־ = Trim(msh.TextMatrix(msh.Row, GetColNum("ɾ����־")))
    
    If cmdModify.Caption = "�޸�(&M)" Then
        If mlng��ҳID = 0 And Trim(msh.TextMatrix(msh.Row, GetColNum("���"))) = "����" Then
            '����21368 by lesfeng 2010-08-02
            If strɾ����־ = "" Then
                cmdModify.Enabled = True
                cmdDel.Enabled = True
                stbThis.Panels(1).Text = "��ǰ������¼��Ч"
            Else
                cmdModify.Enabled = False
                cmdDel.Enabled = False
                stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
            End If
        Else
            If UBound(Split(Trim(msh.TextMatrix(msh.Row, GetColNum("���"))), "��סԺ")) > 0 Then 'ȡ��ѡ������ҳID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(msh.Row, GetColNum("���"))), "��סԺ")(0), "��")(1))
            Else
                RowPageid = 0
            End If
            If mlng��ҳID <> RowPageid Then
                cmdModify.Enabled = False
                cmdDel.Enabled = False
                stbThis.Panels(1).Text = "��ǰ������¼�Ǳ���סԺ������"
            Else
                str����ʱ�� = Trim(msh.TextMatrix(msh.Row, GetColNum("����ʱ��")))
            
                If str����ʱ�� <> "" Then
                    If CDate(str����ʱ��) < Datsys Then
                         cmdModify.Enabled = False
                         cmdDel.Enabled = False
                        '����21368 by lesfeng 2010-08-02
                         If strɾ����־ = "" Then
                            stbThis.Panels(1).Text = "��ǰ������¼�ѹ���"
                        Else
                            stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
                        End If
                    Else
                        '����21368 by lesfeng 2010-08-02
                        If strɾ����־ = "" Then
                            cmdModify.Enabled = True
                            cmdDel.Enabled = True
                            stbThis.Panels(1).Text = "��ǰ������¼��Ч"
                        Else
                            cmdModify.Enabled = False
                            cmdDel.Enabled = False
                            stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
                        End If
                    End If
                Else
                    '����21368 by lesfeng 2010-08-02
                    If strɾ����־ = "" Then
                        cmdModify.Enabled = True
                        cmdDel.Enabled = True
                        stbThis.Panels(1).Text = "��ǰ������¼��Ч"
                    Else
                        cmdModify.Enabled = False
                        cmdDel.Enabled = False
                        stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    Dim blnCard As Boolean, blnICCard As Boolean
    Dim dblMoney As Double, lngRow As Long
    
    If txtPatient.Locked Then Exit Sub
        
    If IDKind.GetCurCard.���� Like "����*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If
    
    If txtPatient.Tag <> "" Then Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        If frmPatiSelect.ShowMe(Me) = False Then
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Exit Sub
        End If
    End If
    Me.Refresh
    mstr�������� = ""
    txtPatient.ForeColor = &HFF0000
    
    'ˢ����ϻ���������س�
    If blnCard And Len(Me.txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtPatient.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        '��ȡ������Ϣ
        Call ClearWinInfor(True)
        
        If IDKind.GetCurCard.���� Like "IC��*" And IDKind.GetCurCard.ϵͳ Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        
        If Not GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), blnCancel, blnCard) Then
            If blnCancel Then 'ȡ������
                Call zlControl.TxtSelAll(txtPatient): txtPatient.SetFocus: Exit Sub
            End If
            stbThis.Panels(1).Text = "δ�ҵ��ò��ˣ�������������!"
            If blnCard = True Then
                txtPatient.PasswordChar = "": txtPatient.Text = "": txtPatient.IMEMode = 0
            Else
                txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
            End If
            Set mrsInfo = New ADODB.Recordset
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Else
            '���ò��˷�����Ϣ
            mlng����ID = Val(Nvl(mrsInfo!����ID, 0))
            mlng��ҳID = Val(Nvl(mrsInfo!��ҳID, 0))
            
            Call ClearWinInfor
            If mrsInfo!��ǰ����id <> 0 Then
                lblBed.Caption = "���ţ�" & IIf(mrsInfo!���� = 0, "��ͥ", mrsInfo!����)
            End If
            
            lblNO.Caption = "סԺ�ţ�" & IIf(mrsInfo!סԺ�� = 0, "", mrsInfo!סԺ��)
            lblDept.Caption = "���ң�" & GET��������(mrsInfo!����ID)
            
            lblType.Caption = "�ѱ�ȼ���" & mrsInfo!�ѱ�
'            lbl������.Caption = lbl������.Tag & mrsInfo!������
'            lbl�������.Caption = lbl�������.Tag & mrsInfo!������
'            chk����temp.Value = mrsInfo!��������
            
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
            txtPatient.Text = mrsInfo!����
            txtPatient.Tag = mrsInfo!����ID
            '-----------------------------------------------------------------------------------------
            lblSex.Caption = "�Ա�" & IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
            lblAge.Caption = "���䣺" & IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
'            lbl��ͥ��ַ.Caption = lbl��ͥ��ַ.Tag & Nvl(mrsInfo!��ͥ��ַ)
            lblCur.Caption = "ҽ�Ƹ��ʽ��" & Nvl(mrsInfo!ҽ�Ƹ��ʽ)
            dblMoney = 0
            For lngRow = 1 To mshList.Rows - 1
                 dblMoney = Format(dblMoney + Val(mshList.TextMatrix(lngRow, GetColNumList("���"))), "#0.00;-#0.00;0.00")
            Next
            lblDeposit.Caption = "Ԥ���ܶ" & IIf(dblMoney = 0, "", dblMoney)
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If mrsInfo Is Nothing Then
            cmdAdd.Enabled = False
        ElseIf mrsInfo.State = adStateClosed Then
            cmdAdd.Enabled = False
        Else
            cmdAdd.Enabled = True
        End If
    End If
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, blnCancel As Boolean, Optional blnCard As Boolean = False) As Boolean
    '���ܣ���ȡ������Ϣ
    '������strInput=[ˢ��]|[A����ID]|[BסԺ��]
    '˵����
    '     �Զ�ʶ������Ժ״̬,����(����ID,��ҳID,����,�Ա�,����,סԺ��,����,��Ժ��־)
    '����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmrsInfo�а���������Ϣ,ʧ��ʱmrsInfo=Close
    Dim rsTmp As ADODB.Recordset, strPati As String, strSQL As String
    Dim vRect As RECT, i As Integer, lng�����ID As Long, bln�����ʻ� As Boolean, lng����ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String, blnICCard As Boolean
    Dim blnHavePassWord As Boolean
    
    blnCancel = False
    strWhere = ""
      
    If (blnCard And objCard.���� Like "����*") _
        And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   'ˢ����ȱʡ�Ŀ�
        lng�����ID = IDKind.GetDefaultCardTypeID
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If mobjSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.����ID=[1]"
        strInput = "-" & lng����ID
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  'סԺ��(��ס(��)Ժ�Ĳ���)
        strWhere = strWhere & " And A.סԺ��=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(�������ﲡ��)
        strWhere = strWhere & " And A.�����=[1]"
    Else '��������
        Select Case objCard.����
            Case "����"
                If Not gblnSeekName Then
                    MsgBox "��ˢ��������[-����ID]��[+סԺ��]��[*�����]�ȷ�ʽ��ȡ���˵���Ϣ��", vbInformation, gstrSysName
                    txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                Else
                    strPati = _
                    " Select A.����ID as ID,A.����ID,C.��ҳID,NVL(C.����,A.����) ����,NVL(C.�Ա�,A.�Ա�) �Ա�,NVL(C.����,A.����) ����," & _
                    "           C.סԺ��,B.���� as ����,A.��ǰ���� as ����," & _
                    "           A.��������,A.���֤��,A.��ͥ��ַ,A.����֤�� " & _
                    " From ������Ϣ A,������ҳ C,���ű� B" & _
                    " Where A.ͣ��ʱ�� is NULL And A.����ID=C.����ID And A.��ҳID=C.��ҳID " & _
                    " And NVL(C.��ҳID,0)<>0 And C.��Ժ���� IS  NULL And A.��ǰ����ID=B.ID(+) And NVL(C.����,A.����) Like [1]" & _
                    "   Order by A.��Ժʱ�� DESC,A.����"
                    vRect = zlControl.GetControlRect(txtPatient.hWnd)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%")
                    If Not rsTmp Is Nothing Then
                        strInput = rsTmp!����ID
                        strWhere = strWhere & " And A.����ID=[2]"
                    Else
                        Set mrsInfo = New ADODB.Recordset: Exit Function
                    End If
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.ҽ����=[2]"
            Case "IC����"
                strInput = UCase(strInput)
                If mobjSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
                blnICCard = (InStr(1, "-+*.", Left(strInput, 1)) = 0) And objCard.ϵͳ
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.סԺ��=[2]"
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    bln�����ʻ� = objCard.�Ƿ�����ʻ�
                    If mobjSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If mobjSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If

    strSQL = _
    " Select A.����ID,Nvl(C.��ҳID,0) as ��ҳID,Nvl(C.��ǰ����ID,0) as ����ID,Nvl(c.��Ժ����ID,0) as ����ID,Nvl(A.��ǰ����ID,0) as ��ǰ����ID, Nvl(a.��Ժ,0) as ��Ժ," & _
    "           Decode(Nvl(A.��ҳID,0),0,A.ҽ�Ƹ��ʽ,C.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ,C.��������," & _
    "           NVL(C.����,A.����) ����,NVL(C.�Ա�,A.�Ա�) �Ա�,NVL(C.����,A.����) ����,Nvl(C.סԺ��,0) as סԺ��,Nvl(C.��Ժ����,0) as ����,A.��ͥ��ַ,A.����֤��," & _
    "           B.����,B.����,Nvl(B.ҽ����,A.ҽ����) ҽ����,B.����,Nvl(C.�ѱ�,A.�ѱ�) �ѱ�,A.������,A.������,Nvl(A.��������,0) as ��������, C.��ע " & _
    " From ������Ϣ A,ҽ�����˵��� B,������ҳ C,ҽ�����˹����� E " & _
    " Where A.ͣ��ʱ�� is NULL" & _
    "       And A.����ID=C.����ID And Nvl(A.��ҳID,0)=C.��ҳID And NVL(C.��ҳID,0)<>0 ANd C.��Ժ���� IS  NULL " & _
    "       And C.����ID=E.����ID(+) And E.��־(+)=1  " & _
    "       And E.ҽ����=B.ҽ����(+) And E.����=B.����(+) And E.���� = B.����(+) " & strWhere
    
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then
        Set mrsInfo = New ADODB.Recordset: Exit Function
    End If
    
    '��Ҫ��������
    If gblnCheckPass And (blnCard Or blnICCard) Then
        If Not blnHavePassWord Then
            strPassWord = Nvl(mrsInfo!����֤��)
        End If
        If strPassWord <> "" Then
            If zlCommFun.VerifyPassWord(Me, strPassWord, mrsInfo!����, mrsInfo!�Ա�, mrsInfo!����) = False Then
                 Set mrsInfo = New ADODB.Recordset: Exit Function
            End If
        End If
    End If
    GetPatient = True
    Exit Function
errH:
     If ErrCenter() = 1 Then Resume
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
End Function


Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(True)
    txtPatient.Tag = ""
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    '����27379 by lesfeng 2010-01-18
    If mrsInfo.State = 1 Then
        mstr�������� = IIf(IsNull(mrsInfo!��������), "", mrsInfo!��������)
    End If
    If mstr�������� = "" Then
        If mrsInfo.State = 1 Then
            If GetOutPatient(mrsInfo!����ID) Then
                txtPatient.ForeColor = vbRed
            Else
                txtPatient.ForeColor = &HFF0000
            End If
        Else
            txtPatient.ForeColor = &HFF0000
        End If
    Else
        txtPatient.ForeColor = zlDatabase.GetPatiColor(mstr��������, True)
    End If
End Sub

Private Function GetOutPatient(ByVal lngID As Long) As Boolean
'���ܣ��ж����ﲡ���Ƿ�����ҽ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim int���� As Integer
    
    GetOutPatient = False
    On Error GoTo errH
    
    strSQL = _
        "Select ���� " & _
        "from ������Ϣ " & _
        "Where ����id = [1] and rownum <= 1 "

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    If Not rsTmp.EOF Then
        int���� = IIf(IsNull(rsTmp!����), -1, rsTmp!����)
        GetOutPatient = int���� <> -1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '����27554 by lesfeng 2010-01-19 lngTXTProc �޸�ΪglngTXTProc
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtReason_GotFocus()
    zlControl.TxtSelAll txtReason
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    Else
        If InStr("'|?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txtReason, KeyAscii
    End If
End Sub

Private Sub txtReason_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OS.OpenImeByName
End Sub

Private Sub txtWarrantM_GotFocus()
    zlControl.TxtSelAll txtWarrantM
End Sub

Private Sub txtWarrantM_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
        If KeyAscii = vbKeyReturn Then
            chkUnlimit.TabStop = (txtWarrantM.Text = "")
            SendKeys "{Tab}"
        Else
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc(".") And InStr(txtWarrantM.Text, ".") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWarrantM_LostFocus()
    If IsNumeric(txtWarrantM.Text) Then
        If txtWarrantM.Text = "999999999" Then
            stbThis.Panels(1).Text = "�����������ֵ����ֵ��ʾ���޵�����"
            If txtWarrantM.Enabled Then txtWarrantM.SetFocus
        Else
            txtWarrantM.Text = Format(txtWarrantM.Text, "0.00")
        End If
    Else
        txtWarrantM.Text = ""
    End If
    
    Call zlCommFun.OpenIme
End Sub

Private Sub txtWarrantP_GotFocus()
    zlControl.TxtSelAll txtWarrantP
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtWarrantP_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txtWarrantP, KeyAscii
    End If
End Sub

Private Sub txtWarrantP_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OS.OpenImeByName
End Sub

Private Function To_Date(ByVal dat���� As Date) As String
'����:������е����ڴ�����ORACLE��Ҫ�����ڸ�ʽ��
    To_Date = "To_Date('" & Format(dat����, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
   
    'ֻ��:ִ�л��˷�ʱ,�ſ��ܹܽ��㿨��
    If blnClosed Then
       If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.CloseWindows
            Set mobjSquareCard = Nothing
        End If
        Exit Sub
    End If
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        Err = 0: On Error GoTo 0:      Exit Sub
    End If
    
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
End Sub

