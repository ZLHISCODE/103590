VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#1.2#0"; "zlIDKind.ocx"
Begin VB.Form frmIDCard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���￨����"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIDCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt���￨ 
      Height          =   360
      Left            =   4800
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   420
      Left            =   210
      TabIndex        =   16
      Top             =   4320
      Width           =   1500
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   165
      ScaleHeight     =   420
      ScaleWidth      =   3300
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   825
      Width           =   3300
      Begin VB.ComboBox cboNO 
         Height          =   360
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "�ȼ�:F12"
         Top             =   30
         Width           =   1560
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2685
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F8"
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   2715
         TabIndex        =   31
         Top             =   30
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   30
         Top             =   90
         Width           =   720
      End
   End
   Begin VB.PictureBox picFace 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   90
      ScaleHeight     =   2895
      ScaleWidth      =   7290
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1275
      Width           =   7290
      Begin VB.TextBox txtAudi 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4335
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1800
         Width           =   1665
      End
      Begin VB.TextBox txt����Ա 
         Height          =   360
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2355
         Width           =   1575
      End
      Begin VB.ComboBox cboStyle 
         Height          =   360
         Left            =   1185
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1259
         Width           =   1590
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1185
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   712
         Width           =   1575
      End
      Begin VB.CheckBox chkBilling 
         Caption         =   "����"
         Height          =   240
         Left            =   3495
         TabIndex        =   6
         Top             =   765
         Width           =   810
      End
      Begin VB.ComboBox cbo���㷽ʽ 
         Height          =   360
         Left            =   4335
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   705
         Width           =   1695
      End
      Begin VB.TextBox txtPatient 
         Height          =   360
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "�ȼ���F11"
         Top             =   165
         Width           =   1815
      End
      Begin VB.TextBox txtSex 
         Height          =   360
         Left            =   4350
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   165
         Width           =   765
      End
      Begin VB.TextBox txtOld 
         Height          =   360
         Left            =   5775
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   165
         Width           =   1300
      End
      Begin VB.TextBox txtCardNO 
         BackColor       =   &H00EBFFFF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4335
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1260
         Width           =   1665
      End
      Begin VB.TextBox txtPass 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1185
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1806
         Width           =   1590
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   4335
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2355
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd hh:mm"
         Mask            =   "####-##-## ##:##"
         PromptChar      =   "_"
      End
      Begin zlIDKind.IDKind IDKind 
         Height          =   360
         Left            =   1185
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "��ݼ�F4"
         Top             =   165
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         IDKindStr       =   "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;��|���￨|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl���㷽ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���㷽ʽ"
         Height          =   240
         Left            =   3315
         TabIndex        =   35
         Top             =   765
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤����"
         Height          =   240
         Left            =   3315
         TabIndex        =   34
         Top             =   1860
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   390
         TabIndex        =   33
         Top             =   2415
         Width           =   720
      End
      Begin VB.Label lblStyle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʽ"
         Height          =   240
         Left            =   630
         TabIndex        =   28
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   240
         Left            =   630
         TabIndex        =   27
         Top             =   765
         Width           =   480
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   240
         Left            =   3315
         TabIndex        =   26
         Top             =   2415
         Width           =   960
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   630
         TabIndex        =   25
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   240
         Left            =   3810
         TabIndex        =   24
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   5280
         TabIndex        =   23
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lblCardNO 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3720
         TabIndex        =   22
         Top             =   1290
         Width           =   540
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   630
         TabIndex        =   21
         Top             =   1860
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   5775
      TabIndex        =   15
      ToolTipText     =   "�ȼ�:Esc"
      Top             =   4320
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   420
      Left            =   4155
      TabIndex        =   14
      Top             =   4320
      Width           =   1500
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   32
      Top             =   4815
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmIDCard.frx":0442
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8229
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblˢ�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���￨"
      Height          =   240
      Left            =   3840
      TabIndex        =   36
      Top             =   880
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���￨���ŵ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   90
      TabIndex        =   19
      Top             =   210
      Width           =   7170
   End
End
Attribute VB_Name = "frmIDCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''Option Explicit 'Ҫ���������
'''''˵�����ڷ���״̬���е��˿�,����ˢ��(������)�����뵥�ݺ�ȷ��
''''Public mbytInState As Byte '�룺0=����(����������),1-�鿴��¼,2-�˿�(���ݺš�����),ֻ֧�ֵ��ݺ�ȷ��
''''Public mblnViewCancel As Boolean '�룺�Ƿ�鿴���˵���
''''Public mstrInNO As String '�룺Ҫ�鿴��Ҫ�˵ĵ��ݺ�,�Ӳ�����Ϣ�Ǽ��е����˿�ʱΪ��
''''Public mblnNOMoved As Boolean '����ϸʱ��¼��ǰѡ��ĵ����Ƿ����������ݱ���,����������ʱ�������ж�
''''
''''Private mblnUnLoad As Boolean
'''''(�շ����,�շ�ϸĿID,���㵥λ,������ĿID,������Ŀ,�վݷ�Ŀ,ԭ��,�ּ�,�Ƿ���,���ұ�־)
''''Private mrs���￨ As ADODB.Recordset
''''Private mrsInfo As New ADODB.Recordset '���没����Ϣ
''''Private mlng�ſ�ID As Long '�ӹ��ü����þ��￨������ѡ�������ID
''''Private mblnICCard As Boolean 'IC������,Ҫͬʱ��д������Ϣ��IC���ֶ�
''''
''''Private WithEvents mobjIDCard As clsIDCard
''''Private mobjICCard As Object
''''Private Enum IDKinds
''''    C0���� = 0
''''    C1ҽ���� = 1
''''    C2���֤�� = 2
''''    C3IC���� = 3
''''    C4����� = 4
''''    C5���￨ = 5
''''End Enum
''''Private mint�˿�ģʽ As Integer
''''Private mstr�˿���֤ As String
''''
''''Private Sub IDKind_Click()
''''    If IDKind.IDKind = IDKinds.C3IC���� Then
''''        If mobjICCard Is Nothing Then
''''            Set mobjICCard = CreateObject("zlICCard.clsICCard")
''''            Set mobjICCard.gcnOracle = gcnOracle
''''        End If
''''        If Not mobjICCard Is Nothing Then
''''            txtPatient.Text = mobjICCard.Read_Card()
''''            If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub lblCardNO_Click()
''''    If txtCardNO.Enabled = False Or txtCardNO.Locked Then Exit Sub
''''    If mobjICCard Is Nothing Then
''''        Set mobjICCard = CreateObject("zlICCard.clsICCard")
''''        Set mobjICCard.gcnOracle = gcnOracle
''''    End If
''''    If Not mobjICCard Is Nothing Then
''''        txtCardNO.Text = mobjICCard.Read_Card()
''''        If txtCardNO.Text <> "" Then mblnICCard = True
''''    End If
''''End Sub
''''
''''Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
''''                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
''''    Dim lngPreIDKind As Long
''''
''''    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
''''        lngPreIDKind = IDKind.IDKind
''''        IDKind.IDKind = IDKinds.C2���֤��
''''        txtPatient.Text = strID
''''        Call txtPatient_KeyPress(vbKeyReturn)
''''        IDKind.IDKind = lngPreIDKind
''''    End If
''''End Sub
''''
''''Private Sub cboStyle_Click()
''''    If Me.Visible Then
''''        If cboStyle.ListIndex = 2 Then '�������,����ȡ���˾��￨����
''''            txtMoney.Text = "0.00"
''''            txtMoney.Locked = True
''''        ElseIf Val(txtMoney.Text) = 0 Then
''''            If mrs���￨!�Ƿ��� = 1 And cboStyle.ListIndex = 0 Then  '����������޼�
''''               txtMoney.Text = Format(mrs���￨!ȱʡ�۸�, "0.00")
''''            Else
''''                txtMoney.Text = Format(mrs���￨!�ּ�, "0.00")
''''            End If
''''            txtMoney.Locked = Not (mrs���￨!�Ƿ��� = 1)
''''            txtMoney.TabStop = (mrs���￨!�Ƿ��� = 1)
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub cboStyle_KeyPress(KeyAscii As Integer)
''''    Dim lngIdx As Long
''''    If KeyAscii = 13 And cboStyle.ListIndex <> -1 Then
''''        Call zlCommFun.PressKey(vbKeyTab)
''''    ElseIf KeyAscii = 13 And cboStyle.ListIndex = -1 Then
''''        Beep
''''    End If
''''    If cboStyle.Locked Then Exit Sub
''''    If SendMessage(cboStyle.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
''''    lngIdx = MatchIndex(cboStyle.hwnd, KeyAscii)
''''    If lngIdx <> -2 Then cboStyle.ListIndex = lngIdx
''''End Sub
''''
''''Private Sub cbo���㷽ʽ_KeyPress(KeyAscii As Integer)
''''    Dim lngIdx As Long
''''    If KeyAscii = 13 And cbo���㷽ʽ.ListIndex <> -1 Then
''''        Call zlCommFun.PressKey(vbKeyTab)
''''    ElseIf KeyAscii = 13 And cbo���㷽ʽ.ListIndex = -1 Then
''''        Beep
''''    End If
''''    If cbo���㷽ʽ.Locked Then Exit Sub
''''    If SendMessage(cbo���㷽ʽ.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
''''    lngIdx = MatchIndex(cbo���㷽ʽ.hwnd, KeyAscii)
''''    If lngIdx <> -2 Then cbo���㷽ʽ.ListIndex = lngIdx
''''End Sub
''''
''''Private Sub chkBilling_Click()
''''    If chkBilling.Value = Checked Then
''''        cbo���㷽ʽ.Enabled = False
''''    ElseIf cbo���㷽ʽ.ListCount = 0 Then
''''        chkBilling.Value = Checked
''''        cbo���㷽ʽ.Enabled = False
''''    Else
''''        cbo���㷽ʽ.Enabled = True
''''    End If
''''    If Visible And txtPatient.Text <> "" And txtCardNO.Enabled Then Call zlCommFun.PressKey(vbKeyTab)
''''End Sub
''''
''''Private Sub chkBilling_KeyPress(KeyAscii As Integer)
''''    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
''''End Sub
''''
''''Private Sub chkCancel_Click()
''''    sta.Panels(2).Text = ""
''''    If chkCancel.Value = Checked Then
''''        '����
''''        chkCancel.ForeColor = &HFF&
''''        If mbytInState = 0 Then
''''            Set txtPatient.Container = Me
''''            txtPatient.Top = picFace.Top + txtSex.Top
''''            txtPatient.Left = picFace.Left + txtMoney.Left + IDKind.Width
''''            txtPatient.PasswordChar = "*"
''''            IDKind.Enabled = False
''''            txtPatient.Locked = True
''''        End If
''''        picFace.Enabled = False
''''        '�����ؽ��������
''''        Call NewCard
''''        txtPatient.Text = ""
''''        '�������˿�ĵ��ݺ�
''''        cboNO.Text = "": cboNO.Tag = ""
''''        cboNO.Locked = False
''''        If cboNO.Visible Then cboNO.SetFocus
''''
''''        '����28130��27929 by lesfeng 2010-02-26 �������˿� ��ʾ
''''        If mbytInState = 0 Then
''''            If mint�˿�ģʽ = 1 Or mint�˿�ģʽ = 3 Then
''''                lblˢ��.Visible = True
''''                txt���￨.Visible = True
''''                If txt���￨.Visible Then txt���￨.SetFocus
''''            End If
''''        End If
''''    Else
''''        '����
''''        chkCancel.ForeColor = 0
''''        If mbytInState = 0 Then
''''            Set txtPatient.Container = picFace
''''            txtPatient.Top = txtSex.Top
''''            txtPatient.Left = txtMoney.Left + IDKind.Width
''''            txtPatient.PasswordChar = ""
''''            IDKind.Enabled = True
''''            txtPatient.Locked = False
''''        End If
''''        picFace.Enabled = True
''''        '�����ؽ��������
''''        Call NewCard
''''        txtPatient.Text = ""
''''        txtMoney.Text = Format(IIf(mrs���￨!�Ƿ��� = 1, mrs���￨!ȱʡ�۸�, mrs���￨!�ּ�), "0.00")
''''        txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm")
''''        '�µ�һ�Ž��ʵ�
''''        cboNO.Locked = True
''''        txtPatient.SetFocus
''''
''''        '����28130��27929 by lesfeng 2010-02-26 �������˿� ����
''''        If mbytInState = 0 Then
''''            If mint�˿�ģʽ = 1 Or mint�˿�ģʽ = 3 Then
''''                lblˢ��.Visible = False
''''                txt���￨.Visible = False
''''                txt���￨.Text = ""
''''            End If
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub cmdCancel_Click()
''''    If mbytInState = 0 And gblnOK Then
''''        If chkCancel.Value = Checked Then
''''            If MsgBox("ȷʵҪ�����˿��˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
''''        Else
''''            If mrsInfo.State = adStateOpen Then
''''                If glngSys Like "8??" Then
''''                    If MsgBox("�ÿͻ��Ļ�Ա����δ����,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
''''                Else
''''                    If MsgBox("�ò��˵ľ��￨��δ����,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
''''                End If
''''            Else
''''                If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
''''            End If
''''        End If
''''    End If
''''    Unload Me
''''End Sub
''''
''''Private Sub cmdHelp_Click()
''''ShowHelp App.ProductName, Me.hwnd, Me.Name
''''End Sub
''''
''''Private Sub cmdOK_Click()
''''    Dim strNO As String, strSQL As String, strCard As String, strICCard As String
''''    Dim i As Integer, strTmp As String
''''    Dim str��֤���� As String
''''    Dim blnTrans As Boolean
''''
''''    If chkCancel.Value = Checked Then
''''        '�˿�
''''        If cboNO.Tag = "" Then
''''            If glngSys Like "8??" Then
''''                MsgBox "�û�Ա�����ż�¼δ��ȷ��ȡ,�����˿���", vbExclamation, gstrSysName
''''                '����31345 by lesfeng 2010-07-08
''''                Exit Sub
''''            Else
''''                MsgBox "�þ��￨���ż�¼δ��ȷ��ȡ,�����˿���", vbExclamation, gstrSysName
''''                '����31345 by lesfeng 2010-07-08
''''                Exit Sub
''''            End If
''''            '����28130��27929 by lesfeng 2010-02-26 �˿���֤����
''''            If mint�˿�ģʽ = 0 Then
''''                '����31345 by lesfeng 2010-07-08
''''                If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus: Exit Sub
''''            Else
''''                '����31345 by lesfeng 2010-07-08
''''                If txt���￨.Enabled And txt���￨.Visible Then txt���￨.SetFocus: Exit Sub
''''            End If
''''        End If
''''        '����28130��27929 by lesfeng 2010-02-26 �˿���֤����
'''''        If mint�˿�ģʽ <> 0 Then
''''        If (mint�˿�ģʽ = 2 Or mint�˿�ģʽ = 3) And txt���￨.Visible Then
''''            str��֤���� = Trim(txtCardNO.Text)
''''            If mstr�˿���֤ = "" Or str��֤���� <> mstr�˿���֤ Then
''''                MsgBox "�˿���֤ʧ�ܣ���˶�ʵ�ʿ����뵱ǰ���ݿ����Ƿ�һ�£�", vbExclamation, gstrSysName
''''                If txt���￨.Enabled And txt���￨.Visible Then txt���￨.SetFocus
''''                Exit Sub
''''            End If
''''        End If
''''
''''        If Not CancelBill(cboNO.Tag) Then '�˿�(��cboNO.Tag=NO)
''''            MsgBox "�˿�����ʧ�ܣ������Ըò�����", vbExclamation, gstrSysName
''''            Exit Sub
''''        End If
''''
''''        mstr�˿���֤ = ""
''''
''''        If mbytInState <> 2 Then
''''            chkCancel.Value = Unchecked '(�������¼�)
''''        Else
''''            gblnOK = True
''''            Unload Me: Exit Sub '�˿�ģʽ�������˳�
''''        End If
''''    Else
''''        '�¾��￨���ż�¼
''''        If mrsInfo.State = adStateClosed Then
''''            If glngSys Like "8??" Then
''''                MsgBox "û��ȷ��Ҫ���Ż�Ա���Ŀͻ�,�ò������ܼ�����", vbExclamation, gstrSysName
''''            Else
''''                MsgBox "û��ȷ��Ҫ���ž��￨�Ĳ���,�ò������ܼ�����", vbExclamation, gstrSysName
''''            End If
''''            txtPatient.SetFocus: Exit Sub
''''        End If
''''
''''        If chkBilling.Value = 0 Then
''''            If cbo���㷽ʽ.ListCount = 0 Then
''''                MsgBox "û�п�ѡ���㷽ʽ,��ʹ�ü��ʷ������ȵ����㷽ʽ�����н������ã�", vbExclamation, gstrSysName
''''                cbo���㷽ʽ.SetFocus: Exit Sub
''''            ElseIf cbo���㷽ʽ.ListIndex = -1 Then
''''                MsgBox "�����ʷ�������ȷ�����㷽ʽ��", vbExclamation, gstrSysName
''''                cbo���㷽ʽ.SetFocus: Exit Sub
''''            End If
''''        End If
''''
''''        '�����:   cboStyle.ListIndex <> 2:����:25930
''''        If mrs���￨!�Ƿ��� = 1 And cboStyle.ListIndex <> 2 Then
''''            If mrs���￨!�ּ� <> 0 And Abs(CCur(txtMoney.Text)) > Abs(mrs���￨!�ּ�) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "��Ա��������ֵ���ܸ�������޼�:" & Format(Abs(mrs���￨!�ּ�), "0.00"), vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "���￨������ֵ���ܸ�������޼�:" & Format(Abs(mrs���￨!�ּ�), "0.00"), vbExclamation, gstrSysName
''''                End If
''''                txtMoney.SetFocus: Exit Sub
''''            End If
''''            If mrs���￨!ԭ�� <> 0 And Abs(CCur(txtMoney.Text)) < Abs(mrs���￨!ԭ��) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "��Ա��������ֵ���ܵ�������޼�:" & Format(Abs(mrs���￨!ԭ��), "0.00"), vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "���￨������ֵ���ܵ�������޼�:" & Format(Abs(mrs���￨!ԭ��), "0.00"), vbExclamation, gstrSysName
''''                End If
''''                txtMoney.SetFocus: Exit Sub
''''            End If
''''        End If
''''
''''        '�������ͼ��
''''        If cboStyle.ListIndex = -1 Then
''''            MsgBox "��ȷ���������ͣ�", vbExclamation, gstrSysName
''''            cboStyle.SetFocus: Exit Sub
''''        End If
''''        If IsNull(mrsInfo!���￨��) Then
''''            If cboStyle.ListIndex <> 0 Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "�޻�Ա���Ŀͻ�ֻ��ѡ�񷢿���", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "�޿�����ֻ��ѡ�񷢿���", vbExclamation, gstrSysName
''''                End If
''''                cboStyle.SetFocus: Exit Sub
''''            End If
''''        Else
''''            If cboStyle.ListIndex = 0 Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "�ֻ�Ա���Ŀͻ�ֻ��ѡ�񲹿��򻻿���", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "�ֿ�����ֻ��ѡ�񲹿��򻻿���", vbExclamation, gstrSysName
''''                End If
''''                cboStyle.SetFocus: Exit Sub
''''            End If
''''        End If
''''
''''        If Not IsDate(txtDate.Text) Then
''''            MsgBox "��������ȷ�ķ���ʱ�䣡", vbExclamation, gstrSysName
''''            txtDate.SetFocus: Exit Sub
''''        End If
''''        If txtCardNO.Text = "" Then
''''           MsgBox "��ˢ��ȷ�����ţ�", vbExclamation, gstrSysName
''''           txtCardNO.SetFocus: Exit Sub
''''        End If
''''
''''        If txtPass.Text <> txtAudi.Text Then
''''            MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
''''            txtPass.Text = "": txtAudi.Text = ""
''''            txtPass.SetFocus: Exit Sub
''''        End If
''''
''''        '����ǰ�����￨�Ƿ��У��Ƿ��ڷ�Χ��
''''        If gblnBill�ſ� Then
''''            mlng�ſ�ID = CheckUsedBill(5, IIf(mlng�ſ�ID > 0, mlng�ſ�ID, glng�ſ�ID), UCase(txtCardNO.Text))
''''            If mlng�ſ�ID <= 0 Then
''''                Select Case mlng�ſ�ID
''''                    Case 0 '����ʧ��
''''                    Case -1
''''                        If glngSys Like "8??" Then
''''                            MsgBox "����û�����ü����õĻ�Ա��,�����ڱ������ù������λ�����һ����Ա����", vbExclamation, gstrSysName
''''                        Else
''''                            MsgBox "����û�����ü����õľ��￨,�����ڱ������ù������λ�����һ�����￨��", vbExclamation, gstrSysName
''''                        End If
''''                    Case -2
''''                        If glngSys Like "8??" Then
''''                            MsgBox "���ع��õĻ�Ա��������,���������ñ��ع��û�Ա�����λ�����һ����Ա����", vbExclamation, gstrSysName
''''                        Else
''''                            MsgBox "���ع��õľ��￨������,���������ñ��ع��þ��￨���λ�����һ�����￨��", vbExclamation, gstrSysName
''''                        End If
''''                    Case -3
''''                        If glngSys Like "8??" Then
''''                            MsgBox "���Ż�Ա���Ų�����Ч��Χ��,�����Ƿ���ȷˢ����", vbExclamation, gstrSysName
''''                        Else
''''                            MsgBox "���ž��￨�Ų�����Ч��Χ��,�����Ƿ���ȷˢ����", vbExclamation, gstrSysName
''''                        End If
''''                        txtCardNO.SetFocus
''''                End Select
''''                Exit Sub
''''            End If
''''        End If
''''
''''        '����
''''        If CByte(cboStyle.ListIndex) = 2 Then
''''            '����,�൱���ش�,����������,��Ҫ��ȡԭ���ݺ�
''''            strNO = GetNOFromCard(mrsInfo!���￨��)
''''            If strNO = "" Then
''''                MsgBox "û�з��ָò�����ǰ�ķ�����¼�����ܲ�����", vbExclamation, gstrSysName
''''                Exit Sub
''''            End If
''''        Else
''''            '����,�����²�������
''''            strNO = zldatabase.GetNextNo(16)
''''        End If
''''
''''        strCard = UCase(txtCardNO.Text)
''''        strICCard = IIf(mblnICCard, strCard, "")
''''
''''        strSQL = SaveIDCard(CByte(cboStyle.ListIndex), strNO, mrsInfo!����ID, mrsInfo!��ҳID, _
''''            IIf(mrsInfo!����ID = 0, UserInfo.����ID, mrsInfo!����ID), _
''''            IIf(mrsInfo!����ID = 0, UserInfo.����ID, mrsInfo!����ID), _
''''            IIf(mrsInfo!סԺ�� = 0, mrsInfo!�����, mrsInfo!סԺ��), IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�), IIf(IsNull(mrsInfo!���￨��), "", mrsInfo!���￨��), _
''''            mrsInfo!����, IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�), IIf(IsNull(mrsInfo!����), "", mrsInfo!����), _
''''            strCard, txtPass.Text, IIf(mrs���￨!�Ƿ��� = 0, mrs���￨!�ּ�, CCur(txtMoney.Text)), CCur(txtMoney.Text), IIf(Not cbo���㷽ʽ.Enabled, "", NeedName(cbo���㷽ʽ.Text)), _
''''            CDate(txtDate.Text), mlng�ſ�ID, mrs���￨, strICCard)
''''
''''        On Error GoTo errH
''''        gcnOracle.BeginTrans: blnTrans = True
'''''        Call SQLTest(App.ProductName, Me.Caption, strSQL)
'''''        gcnOracle.Execute strSQL, , adCmdStoredProc
'''''        Call SQLTest
''''        zldatabase.ExecuteProcedure strSQL, Me.Caption
''''        gcnOracle.CommitTrans: blnTrans = False
''''
''''
''''        On Error GoTo 0
''''        '���˺�:24662
''''        Dim strOutPut As String
''''        'If Not mobjICCard Is Nothing Then
''''        Call zlExcuteUploadSwap(Val(Nvl(mrsInfo!����ID)), strOutPut, mobjICCard)
''''        'End If
''''
''''        '��ӡ(�Ǽ���,�ǻ���ʱ)
''''
''''        '���뵥����ʷ��¼(�������͵���)
''''        For i = 0 To cboNO.ListCount - 1
''''            strTmp = strTmp & "," & cboNO.List(i)
''''        Next
''''        strTmp = strNO & strTmp
''''        cboNO.Clear
''''        For i = 0 To UBound(Split(strTmp, ","))
''''            cboNO.AddItem Split(strTmp, ",")(i)
''''            If i = 9 Then Exit For 'ֻ��ʾ10��
''''        Next
''''    End If
''''
''''    gblnOK = True
''''
''''    '�����ؽ��������
''''    Call NewCard(False)
''''    mblnICCard = False  '���ܷ���newcard��,��Ϊ�����ȶ����ٶ�����
''''    txtPatient.Text = ""
''''    If Val(txtMoney.Text) = 0 Then txtMoney.Text = Format(mrs���￨!�ּ�, "0.00")
''''    txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm")
''''
''''    '���������￨�Ƿ���
''''    '���Զ������¿���(��ͬ��Ʊ��)
''''    If gblnBill�ſ� Then
''''        mlng�ſ�ID = CheckUsedBill(5, IIf(mlng�ſ�ID > 0, mlng�ſ�ID, glng�ſ�ID))
''''        If mlng�ſ�ID <= 0 Then
''''            Select Case mlng�ſ�ID
''''                Case 0 '����ʧ��
''''                Case -1
''''                    If glngSys Like "8??" Then
''''                        MsgBox "����û�����ü����õĻ�Ա��,�����ڱ������ù������λ�����һ����Ա����", vbExclamation, gstrSysName
''''                    Else
''''                        MsgBox "����û�����ü����õľ��￨,�����ڱ������ù������λ�����һ�����￨��", vbExclamation, gstrSysName
''''                    End If
''''                Case -2
''''                    If glngSys Like "8??" Then
''''                        MsgBox "���ع��õĻ�Ա��������,���������ñ��ع��û�Ա�����λ�����һ����Ա����", vbExclamation, gstrSysName
''''                    Else
''''                        MsgBox "���ع��õľ��￨������,���������ñ��ع��þ��￨���λ�����һ�����￨��", vbExclamation, gstrSysName
''''                    End If
''''            End Select
''''            Exit Sub
''''        End If
''''    End If
''''
''''    txtPatient.SetFocus
''''    Exit Sub
''''errH:
''''    If blnTrans Then gcnOracle.RollbackTrans
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''End Sub
''''
''''Private Sub Form_Activate()
''''    If mbytInState = 2 Then
''''        '����28130��27929 by lesfeng 2010-02-26 �������˿� ��ʾ
''''        If mint�˿�ģʽ = 2 Or mint�˿�ģʽ = 3 Then
''''            lblˢ��.Visible = True
''''            txt���￨.Visible = True
''''            If txt���￨.Visible Then txt���￨.SetFocus
''''        Else
''''            cmdOK.SetFocus
''''        End If
''''    ElseIf mbytInState = 1 Then
''''        cmdCancel.SetFocus
''''    End If
''''End Sub
''''
''''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''''    On Error Resume Next
''''    Select Case KeyCode
''''        Case vbKeyF4
''''            If Shift = vbCtrlMask Then
''''                If IDKind.Enabled Then IDKind.IDKind = IDKinds.C3IC����: Call IDKind_Click
''''            ElseIf Me.ActiveControl Is txtPatient Then
''''                If IDKind.Enabled Then
''''                    If Shift = vbShiftMask Then
''''                        IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDKindStr, ";")), IDKind.IDKind - 1)
''''                    Else
''''                        IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDKindStr, ";")), 0, IDKind.IDKind + 1)
''''                    End If
''''                End If
''''            End If
''''        Case vbKeyF8
''''            If chkCancel.Visible And picNo.Enabled Then
''''                chkCancel.Value = IIf(chkCancel.Value = Checked, Unchecked, Checked)
''''            End If
''''        Case vbKeyF11
''''            If Not txtPatient.Locked Then txtPatient.SetFocus
''''        Case vbKeyF12
''''            If Not cboNO.Locked And picNo.Enabled Then cboNO.SetFocus
''''        Case vbKeyEscape
''''            cmdCancel_Click
''''    End Select
''''End Sub
''''
''''Private Sub Form_KeyPress(KeyAscii As Integer)
''''    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
''''End Sub
''''
''''Private Sub Form_Load()
''''    If glngSys Like "8??" Then
''''        Caption = "��Ա������"
''''        lblPatient.Caption = "�ͻ�"
''''        lblTitle.Caption = gstrUnitName & "��Ա�����ŵ�"
''''        chkBilling.Visible = False
''''        lbl���㷽ʽ.Visible = True
''''    Else
''''        lblTitle.Caption = gstrUnitName & "���￨���ŵ�"
''''    End If
''''
''''    mblnUnLoad = False
''''    '����28130��27929 by lesfeng 2010-02-26
''''    mint�˿�ģʽ = Val(zldatabase.GetPara("�˿�ˢ��", glngSys, glngModul))
'''''    mint�˿�ģʽ = 3
''''    mstr�˿���֤ = ""
''''    '���￨���ü��
''''    If mbytInState = 0 And gblnBill�ſ� Then
''''        mlng�ſ�ID = CheckUsedBill(5, IIf(mlng�ſ�ID > 0, mlng�ſ�ID, glng�ſ�ID))
''''        If mlng�ſ�ID <= 0 Then
''''            Select Case mlng�ſ�ID
''''                Case 0 '����ʧ��
''''                Case -1
''''                    If glngSys Like "8??" Then
''''                        MsgBox "����û�����ü����õĻ�Ա��,�����ڱ������ù������λ�����һ����Ա����", vbExclamation, gstrSysName
''''                    Else
''''                        MsgBox "����û�����ü����õľ��￨,�����ڱ������ù������λ�����һ�����￨��", vbExclamation, gstrSysName
''''                    End If
''''                Case -2
''''                    If glngSys Like "8??" Then
''''                        MsgBox "���ع��õĻ�Ա��������,���������ñ��ع��û�Ա�����λ�����һ����Ա����", vbExclamation, gstrSysName
''''                    Else
''''                        MsgBox "���ع��õľ��￨������,���������ñ��ع��þ��￨���λ�����һ�����￨��", vbExclamation, gstrSysName
''''                    End If
''''            End Select
''''            mblnUnLoad = True: Unload Me: Exit Sub
''''        End If
''''    End If
''''
''''    RestoreWinState Me
''''
''''    Call InitFace
''''    If mblnUnLoad Then: Unload Me: Exit Sub
''''
''''    Call RaisEffect(picFace, -1)
''''End Sub
''''
''''Private Sub Form_Unload(Cancel As Integer)
''''    mblnICCard = False
''''    mbytInState = 0
''''    mblnViewCancel = False
''''    mstrInNO = ""
''''    mlng�ſ�ID = 0
''''    mblnUnLoad = False
''''    mblnNOMoved = False
''''    Set mobjICCard = Nothing
''''    If Not mobjIDCard Is Nothing Then
''''        Call mobjIDCard.SetEnabled(False)
''''        Set mobjIDCard = Nothing
''''    End If
''''End Sub
''''
''''Private Sub InitFace()
''''    Dim rsTmp As New ADODB.Recordset
''''    Dim i As Integer, strSQL As String
''''    gblnOK = False
''''
''''    If gblnShowCard Then txtCardNO.PasswordChar = ""
''''
''''    IDKind.Enabled = mbytInState = 0
''''
''''    Select Case mbytInState
''''        Case 0 '����
''''            Set mrsInfo = New ADODB.Recordset
''''            Set mobjIDCard = New clsIDCard
''''
''''            txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm")
''''
''''            cboStyle.AddItem "1-����"
''''            cboStyle.AddItem "2-����"
''''            cboStyle.AddItem "3-����" '����ȡ���,���������,������Ʊ��
''''            cboStyle.ListIndex = 0
''''            chkBilling.Value = IIf(gbln���� = True, 1, 0)
''''            txt����Ա.Text = UserInfo.����
''''
''''            '���㷽ʽ
''''            strSQL = _
''''                "Select B.����,B.����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
''''                " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
''''                " Where A.Ӧ�ó���='���￨' And B.����=A.���㷽ʽ And Nvl(B.����,1) IN(1,2)" & _
''''                " Order by B.����"
''''            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
''''
''''            If Not rsTmp.EOF Then
''''                For i = 1 To rsTmp.RecordCount
''''                    cbo���㷽ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
''''                    If rsTmp!ȱʡ = 1 Then
''''                        cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
''''                        cbo���㷽ʽ.ItemData(cbo���㷽ʽ.NewIndex) = 1
''''                    End If
''''                    rsTmp.MoveNext
''''                Next
''''                If cbo���㷽ʽ.ListIndex = -1 Then cbo���㷽ʽ.ListIndex = 0
''''            Else
''''                '�޽��㷽ʽֻ�ܼ��ʷ���
''''                If glngSys Like "8??" Then
''''                    MsgBox "��Ա������û�п��õĽ��㷽ʽ�����ȵ����㷽ʽ���������á�", vbInformation, gstrSysName
''''                    mblnUnLoad = True: Exit Sub
''''                Else
''''                    MsgBox "���￨����û�п��õĽ��㷽ʽ��ֻ��ʹ�ü��ʷ�ʽ������", vbInformation, gstrSysName
''''                End If
''''                chkBilling.Value = 1
''''                chkBilling.Enabled = False
''''                cbo���㷽ʽ.Enabled = False
''''            End If
''''
''''            '���￨����
''''            Set mrs���￨ = GetSpecialInfo("���￨")
''''            If Not mrs���￨ Is Nothing Then
''''                If Not mrs���￨.EOF Then
''''                    txtMoney.Locked = Not (mrs���￨!�Ƿ��� = 1)
''''                    txtMoney.TabStop = (mrs���￨!�Ƿ��� = 1)
''''                    If mrs���￨!�Ƿ��� = 1 Then
''''                        txtMoney.Text = Format(mrs���￨!ȱʡ�۸�, "0.00")
''''                    Else
''''                        txtMoney.Text = Format(mrs���￨!�ּ�, "0.00")
''''                    End If
''''                End If
''''            Else
''''                If glngSys Like "8??" Then
''''                    MsgBox "��δ���û�Ա���շ���Ϣ�����ȵ�ҩ�����в��������ã�", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "��δ���þ��￨�շ���Ϣ�����ȵ��������������д���", vbExclamation, gstrSysName
''''                End If
''''                mblnUnLoad = True: Exit Sub
''''            End If
''''        Case 1 'Ԥ��
''''            chkCancel.Visible = False
''''            If mblnViewCancel Then lblFlag.Visible = True
''''            picNo.Enabled = False
''''            picFace.Enabled = False
''''            cmdOK.Visible = False
''''
''''            cmdCancel.Caption = "�˳�(&X)"
''''
''''            If Not ReadBill(mstrInNO) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "������ȷ��ȡ�û�Ա�����ż�¼��", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "������ȷ��ȡ�þ��￨���ż�¼��", vbExclamation, gstrSysName
''''                End If
''''                mblnUnLoad = True: Exit Sub
''''            End If
''''        Case 2  '�˿�
''''            chkCancel.Value = Checked 'ͬʱ�����¼�
''''            picFace.Enabled = False
''''            picNo.Enabled = False
''''
''''            If Not ReadBill(mstrInNO) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "������ȷ��ȡ�û�Ա�����ż�¼��", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "������ȷ��ȡ�þ��￨���ż�¼��", vbExclamation, gstrSysName
''''                End If
''''                mblnUnLoad = True: Exit Sub
''''            End If
''''            '����28130��27929 by lesfeng 2010-02-26
''''            If mint�˿�ģʽ = 2 Or mint�˿�ģʽ = 3 Then
''''                lblˢ��.Visible = True
''''                txt���￨.Visible = True
''''                If txt���￨.Visible Then txt���￨.SetFocus
''''            End If
''''
''''    End Select
''''End Sub
''''
''''Private Sub txtAudi_GotFocus()
''''    SelAll txtAudi
''''    If glngSys Like "8??" Then
''''        sta.Panels(2) = "��ͻ��ٴ�������ͬ�ĵ����룡"
''''    Else
''''        sta.Panels(2) = "�벡���ٴ�������ͬ�ĵ����룡"
''''    End If
''''End Sub
''''
''''Private Sub txtAudi_KeyPress(KeyAscii As Integer)
''''    If KeyAscii = 13 Then
''''        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
''''    Else
''''        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
''''    End If
''''End Sub
''''
''''Private Sub txtCardNO_GotFocus()
''''    SelAll txtCardNO
''''    sta.Panels(2) = "�뽫�ſ���ˢ���������Ữ����"
''''    Call Beep: Beep
''''End Sub
''''
''''Private Sub txtCardNO_KeyPress(KeyAscii As Integer)
'''''    If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
''''    KeyAscii = Asc(UCase(Chr(KeyAscii)))
''''    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
''''
''''    If KeyAscii <> 13 Then
''''        If Len(txtCardNO.Text) = gbytCardNOLen - 1 And KeyAscii <> 8 Then
''''            txtCardNO.Text = txtCardNO.Text & Chr(KeyAscii)
''''            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
''''        End If
''''    Else
''''        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
''''    End If
''''End Sub
''''
''''Private Sub txtCardNO_LostFocus()
''''    sta.Panels(2) = ""
''''End Sub
''''
''''Private Sub txtDate_GotFocus()
''''    txtDate.SelStart = 8
''''    txtDate.SelLength = Len(txtDate.Text) - 8
''''End Sub
''''
''''Private Sub txtDate_KeyPress(KeyAscii As Integer)
''''    If IsDate(txtDate.Text) And KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
''''End Sub
''''
''''Private Sub txtMoney_GotFocus()
''''    If Not txtMoney.Locked Then SelAll txtMoney
''''End Sub
''''
''''Private Sub txtMoney_KeyPress(KeyAscii As Integer)
''''    If KeyAscii = 13 And txtMoney.Text <> "" Then
''''        Call zlCommFun.PressKey(vbKeyTab)
''''    Else
''''        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
''''    End If
''''End Sub
''''
''''Private Sub txtMoney_Validate(Cancel As Boolean)
''''    If IsNumeric(txtMoney.Text) Then txtMoney.Text = Format(txtMoney.Text, "0.00")
''''    '�����
''''    If mbytInState = 0 Then
''''        If cboStyle.ListIndex = 2 Then Exit Sub
''''        If mrs���￨!�Ƿ��� = 1 Then
''''            If mrs���￨!�ּ� <> 0 And Abs(CCur(txtMoney.Text)) > Abs(mrs���￨!�ּ�) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "��Ա��������ֵ���ܸ�������޼�:" & Format(Abs(mrs���￨!�ּ�), "0.00"), vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "���￨������ֵ���ܸ�������޼�:" & Format(Abs(mrs���￨!�ּ�), "0.00"), vbExclamation, gstrSysName
''''                End If
''''                SelAll txtMoney: Cancel = True: Exit Sub
''''            End If
''''
''''            If mrs���￨!ԭ�� <> 0 And Abs(CCur(txtMoney.Text)) < Abs(mrs���￨!ԭ��) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "��Ա��������ֵ���ܵ�������޼�:" & Format(Abs(mrs���￨!ԭ��), "0.00"), vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "���￨������ֵ���ܵ�������޼�:" & Format(Abs(mrs���￨!ԭ��), "0.00"), vbExclamation, gstrSysName
''''                End If
''''                SelAll txtMoney: Cancel = True: Exit Sub
''''            End If
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub txtPass_GotFocus()
''''    SelAll txtPass
''''    If glngSys Like "8??" Then
''''        sta.Panels(2) = "��ͻ�����10λ���ڵ����룡"
''''    Else
''''        sta.Panels(2) = "�벡������10λ���ڵ����룡"
''''    End If
''''End Sub
''''
''''Private Sub txtPass_KeyPress(KeyAscii As Integer)
''''    If KeyAscii = 13 Then
''''        KeyAscii = 0
''''        If txtPass.Text = "" And txtAudi.Text = "" Then
''''            cmdOK.SetFocus
''''        Else
''''            Call zlCommFun.PressKey(vbKeyTab)
''''        End If
''''    Else
''''        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
''''    End If
''''End Sub
''''
''''Private Sub txtPass_LostFocus()
''''    sta.Panels(2) = ""
''''End Sub
''''
''''Private Sub cboNO_GotFocus()
''''    If Not cboNO.Locked Then SelAll cboNO
''''End Sub
''''
''''Private Sub cboNO_KeyPress(KeyAscii As Integer)
''''    Dim strOper As String, vDate As Date
''''
''''    If cboNO.Locked Then Exit Sub
''''
''''    'ת���ɴ�д(���ֲ��ɴ���)
''''    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
''''
''''    '��һλ����������ĸ,����λ����
''''    If KeyAscii <> 13 Then
''''        Call SetNOInputLimit(cboNO, KeyAscii)
''''    ElseIf cboNO.Text <> "" And Not cboNO.Locked Then
''''        cboNO.Text = GetFullNO(cboNO.Text, 16)
''''
''''        '�Ƿ���ת������ݱ���
''''        If zldatabase.NOMoved("סԺ���ü�¼", cboNO.Text, , "5") Then
''''            If Not ReturnMovedExes(cboNO.Text, 5, Me.Caption) Then Exit Sub
''''            mblnNOMoved = False
''''        End If
''''
''''        '����Ȩ��
''''        If Not ReadBillInfo(2, cboNO.Text, 5, strOper, vDate) Then
''''            txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Sub
''''        End If
''''        If Not BillOperCheck(8, strOper, vDate, "�˿�") Then
''''            txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Sub
''''        End If
''''
''''        '��ȡҪ�˿��ļ�¼(��NO)
''''        Select Case ReadBill(cboNO.Text)
''''            Case -1
''''                '����28130��27929 by lesfeng 2010-02-26 �������˿� ��ʾ
''''                If mint�˿�ģʽ = 1 Or mint�˿�ģʽ = 3 Then
''''                    If txt���￨.Visible Then txt���￨.SetFocus
''''                Else
''''                    cmdOK.SetFocus
''''                End If
''''            Case 0
''''                If glngSys Like "8??" Then
''''                    MsgBox "��ȡ�û�Ա�����ż�¼ʧ�ܣ�", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "��ȡ�þ��￨���ż�¼ʧ�ܣ�", vbExclamation, gstrSysName
''''                End If
''''                txtPatient.Text = "": cboNO.SetFocus
''''            Case 1
''''                If glngSys Like "8??" Then
''''                    MsgBox "�û�Ա�����ż�¼�����ڣ�", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "�þ��￨���ż�¼�����ڣ�", vbExclamation, gstrSysName
''''                End If
''''                txtPatient.Text = "": cboNO.SetFocus
''''            Case 2
''''                If glngSys Like "8??" Then
''''                    MsgBox "�û�Ա�����ż�¼�Ѿ��˳���", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "�þ��￨���ż�¼�Ѿ��˳���", vbExclamation, gstrSysName
''''                End If
''''                txtPatient.Text = "": cboNO.SetFocus
''''        End Select
''''    End If
''''End Sub
''''
''''Private Sub txtPatient_Change()
''''    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
''''End Sub
''''
''''Private Sub txtPatient_GotFocus()
''''    SelAll txtPatient
''''    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(True)
''''End Sub
''''
''''Private Sub txtPatient_KeyPress(KeyAscii As Integer)
''''    Dim strNO As String
''''    Dim blnCard As Boolean
''''
''''    If txtPatient.Locked Then Exit Sub
''''    '�����ַ�������Form_KeyPress�н���
''''    'ˢ����ʾ����
''''    If chkCancel.Value = Checked Then txtPatient.PasswordChar = IIf(gblnShowCard, "", "*")
''''
''''    '-010��+010ʱIsNumeric(txtPatient.Text)����true,��ʱtxtPatient.Text�ĳ��Ȳ���������������ַ�
''''    If IDKind.IDKind = IDKinds.C0���� Then
''''        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, glngSys)
''''    ElseIf IDKind.IDKind = IDKinds.C4����� Then
''''        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
''''            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
''''        End If
''''    End If
''''
''''    If blnCard And Len(txtPatient.Text) = gbytCardNOLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txtPatient.Text <> "" Then
''''        If KeyAscii <> 13 Then
''''            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
''''            txtPatient.SelStart = Len(txtPatient.Text)
''''        End If
''''        KeyAscii = 0
''''        sta.Panels(2) = ""
''''
''''        If chkCancel.Value = 1 Then
''''            If blnCard = True Then
''''                '1.��ȡҪ�˿��ļ�¼(�ɿ���)
''''                strNO = GetNOFromCard(txtPatient.Text)
''''                If strNO = "" Then
''''                    If glngSys Like "8??" Then
''''                        sta.Panels(2) = "���ܶ�ȡ�ͻ�������Ϣ,��ȷ���Ƿ���ȷˢ����"
''''                    Else
''''                        sta.Panels(2) = "���ܶ�ȡ���˷�����Ϣ,��ȷ���Ƿ���ȷˢ����"
''''                    End If
''''                    txtPatient.Text = "": txtPatient.SetFocus
''''                    Call Beep: Exit Sub
''''                End If
''''                '�����˿�
''''                cboNO.Text = strNO
''''                Call cboNO_KeyPress(13): Exit Sub
''''            Else
''''                sta.Panels(2) = "���ܶ�ȡ�Ϳ���Ϣ,��ȷ���Ƿ���ȷˢ����"
''''                txtPatient.Text = "": txtPatient.SetFocus
''''                Call Beep: Exit Sub
''''            End If
''''        Else
''''            '2.���뿨��,��סԺ�ŵȲ��˱�ʶ,��,��,����
''''            Call NewCard(False)  '�����Ϣ
''''            '��ȡ������Ϣ
''''            If Not GetPatient(txtPatient.Text, blnCard) Then
''''                If glngSys Like "8??" Then
''''                    sta.Panels(2) = "û�з��ָÿͻ���Ϣ,����δ����,���飡"
''''                Else
''''                    sta.Panels(2) = "û�з��ָò�����Ϣ,����δ����,���飡"
''''                End If
''''                txtPatient.Text = "": txtPatient.SetFocus
''''                Call Beep: Exit Sub
''''            End If
''''            txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm")
''''            If cboStyle.ListIndex = 0 And mrs���￨!�Ƿ��� = 1 Then
''''                txtMoney.Text = Format(mrs���￨!ȱʡ�۸�, "0.00")
''''            Else
''''                txtMoney.Text = Format(mrs���￨!�ּ�, "0.00")
''''
''''                If mrs���￨!�Ƿ��� = 0 Then
''''                    txtMoney.Text = Format(GetActualMoney(Nvl(mrsInfo!�ѱ�), mrs���￨!������ĿID, mrs���￨!�ּ�, mrs���￨!�շ�ϸĿID), "0.00")
''''                End If
''''            End If
''''
''''            If Not IsNull(mrsInfo!���￨��) Then
''''                If MsgBox(IIf(glngSys Like "8??", "�ͻ�", "����") & "�Ѿ�����" & IIf(glngSys Like "8??", "��Ա", "����") & "����Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
''''                    Set mrsInfo = New ADODB.Recordset
''''                    txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
''''                    txtPatient.SetFocus: Exit Sub
''''                Else
''''                    If cboStyle.ListCount > 2 And blnCard = True Then  '������
''''                        cboStyle.ListIndex = 2 '����ǻ���,���Զ�����Ϊ����
''''                    Else
''''                        cboStyle.ListIndex = 1 '������䲡�˱�ʶ,���Զ���Ϊ����
''''                    End If
''''                End If
''''            End If
''''            '�����µķ��ż�¼
''''            txtPatient.Text = mrsInfo!����
''''            txtSex.Text = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
''''            txtOld.Text = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
''''            Call zlCommFun.PressKey(vbKeyTab) '��λ���������
''''            Exit Sub
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub NewCard(Optional blnMoney As Boolean = True)
''''    Set mrsInfo = New ADODB.Recordset
''''    cboNO.Text = ""
''''    If blnMoney Then txtMoney.Text = ""
''''    txtSex.Text = ""
''''    txtOld.Text = ""
''''    chkBilling.Value = IIf(gbln���� = True, 1, 0)
''''    txtCardNO.Text = ""
''''    txtPass.Text = ""
''''    txtAudi.Text = ""
''''    txtDate.Text = "____-__-__ __:__"
''''    If cboStyle.ListCount > 0 Then cboStyle.ListIndex = 0
''''End Sub
''''
''''Private Function GetPatient(ByVal strInput As String, Optional blnCard As Boolean = False) As Boolean
'''''���ܣ���ȡ������Ϣ
'''''������strInput=���˱�ʶ��(A-����ID,B+סԺ��,C/����,D*�����,G.�Һŵ���)
'''''����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmrsInfo�а���������Ϣ,ʧ��ʱmrsInfo=Close
''''    Dim strSQL As String, objRect As RECT
''''    On Error GoTo errH
''''
''''    '������Ժʱ��סԺ�ѱ�,����������ѱ�
''''    strSQL = _
''''        " Select Rownum id,A.���￨��,A.����ID,Nvl(B.��ҳID,0) as ��ҳID," & _
''''        " Nvl(A.��ǰ����ID,0) as ����ID,Nvl(A.��ǰ����ID,0) as ����ID," & _
''''        " A.����,A.�Ա�,A.����,Nvl(A.סԺ��,0) as סԺ��," & _
''''        " Nvl(A.�����,0) as �����,Nvl(A.��ǰ����,0) as ����," & _
''''        " Decode(A.��ǰ����ID,NULL,A.�ѱ�,B.�ѱ�) as �ѱ�,A.��ͥ��ַ" & _
''''        " From ������Ϣ A,������ҳ B" & _
''''        " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID(+)" & _
''''        " And Nvl(A.סԺ����,0)=B.��ҳID(+) And Nvl(B.��ҳID(+),0)<>0"
''''    If Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
''''        strSQL = strSQL & " And A.����ID=[1] and '%'='%'"
''''    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(����סԺ����)
''''        strSQL = strSQL & " And A.��ǰ����ID is Not NULL And A.סԺ��=[1] and '%'='%'"
''''    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(�������ﲡ��)
''''        strSQL = strSQL & " And A.��ǰ����ID is NULL And A.�����=[1] and '%'='%'"
''''    ElseIf blnCard = True Or IDKind.IDKind = IDKinds.C5���￨ Then
''''        strInput = UCase(strInput)
''''        strSQL = strSQL & " And A.���￨��=[2] and '%'='%'"  '���￨��
''''    Else '��������
''''        Select Case IDKind.IDKind
''''            Case IDKinds.C0����
''''                strSQL = strSQL & " And A.����=[2] and '%'='%'"
''''            Case IDKinds.C1ҽ����
''''                strInput = UCase(strInput)
''''                strSQL = strSQL & " And A.ҽ����=[2] and '%'='%'"
''''            Case IDKinds.C2���֤��
''''                strInput = UCase(strInput)
''''                strSQL = strSQL & " And A.���֤��=[2] and '%'='%'"
''''            Case IDKinds.C3IC����
''''                strInput = UCase(strInput)
''''                strSQL = strSQL & " And A.IC����=[2] and '%'='%'"
''''            Case IDKinds.C4�����
''''                If Not IsNumeric(strInput) Then strInput = "0"
''''                strSQL = strSQL & " And A.��ǰ����ID is NULL And A.�����=[2] and '%'='%'"
''''        End Select
''''    End If
''''
''''    'Set mrsInfo = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
''''    objRect = GetControlRect(txtPatient.hwnd)
''''    Set mrsInfo = zldatabase.ShowSQLSelect(Me, strSQL, 0, "�鲡����Ϣ", False, "����ID", "", False, False, True, objRect.Left, objRect.Top, txtPatient.Height, False, True, False, Mid(strInput, 2), strInput)
''''
''''    If mrsInfo.State = adStateClosed Then
''''        Set mrsInfo = New ADODB.Recordset
''''    Else
''''        GetPatient = True
''''    End If
''''    Exit Function
''''errH:
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''    Set mrsInfo = New ADODB.Recordset
''''End Function
''''
''''Private Function ReadBill(strNO As String) As Integer
'''''���ܣ��ɵ��ݺŶ�ȡ����ʾ���￨���ż�¼
'''''���أ�
'''''     -1:�ɹ�
'''''      0:ʧ��
'''''      1:�ü�¼������
'''''      2:�ü�¼�Ѿ�����(��mblnViewCancel=Falseʱ��Ч)
''''    On Error GoTo errH
''''
''''    Dim rsTmp As New ADODB.Recordset
''''    Dim strFullNO As String
''''
''''    strFullNO = GetFullNO(strNO, 16)
''''    '��Ϊ���￨���õĽ���ID�����Ǽ��ʷ��������ʱ������ID,
''''    '������Ԥ����¼����ʱһ��Ҫ�Ӽ�¼����=5����
''''    'by lesfeng 2010-03-08 ����A.* ����
''''    gstrSQL = _
''''        " Select A.NO,A.����,A.�Ա�,A.����,A.ʵ��Ʊ��,A.���ӱ�־,A.��¼״̬,A.ʵ�ս��,A.����Ա����,A.����ʱ��,B.����֤��,C.���㷽ʽ " & _
''''        " From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A,������Ϣ B," & _
''''        " (Select ���㷽ʽ,����ID From " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ Where ��¼����=5 And NO=[1]) C" & _
''''        " Where A.����ID=C.����ID(+) And A.��¼����=5 And A.����ID=B.����ID And A.NO=[1] " & _
''''         IIf(mblnViewCancel, "And A.��¼״̬=3 ", "")
''''
''''    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFullNO)
''''    If rsTmp.EOF Then ReadBill = 1: Exit Function
''''
''''    If Not mblnViewCancel And (rsTmp!��¼״̬ = 3 Or rsTmp!��¼״̬ = 2) Then
''''        ReadBill = 2: Exit Function
''''    End If
''''
''''    cboNO.Text = rsTmp!NO
''''    cboNO.Tag = rsTmp!NO
''''    txtPatient.Text = rsTmp!����
''''    txtPatient.PasswordChar = ""
''''    txtSex.Text = IIf(IsNull(rsTmp!�Ա�), "", rsTmp!�Ա�)
''''    txtOld.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
''''
''''    If IsNull(rsTmp!���㷽ʽ) Then
''''        chkBilling.Value = Checked
''''    Else
''''        chkBilling.Value = Unchecked
''''        If cbo���㷽ʽ.ListCount = 0 Then
''''            cbo���㷽ʽ.AddItem rsTmp!���㷽ʽ
''''            cbo���㷽ʽ.ListIndex = 0
''''        Else
''''            cbo���㷽ʽ.ListIndex = GetCboIndex(cbo���㷽ʽ, rsTmp!���㷽ʽ)
''''        End If
''''    End If
''''    txtCardNO.Text = IIf(IsNull(rsTmp!ʵ��Ʊ��), "", rsTmp!ʵ��Ʊ��)
''''    txtPass.Text = IIf(IsNull(rsTmp!����֤��), "", rsTmp!����֤��)
''''    txtAudi.Text = txtPass.Text
''''
''''    If cboStyle.ListCount = 0 Then
''''        Select Case rsTmp!���ӱ�־
''''            Case 0
''''                cboStyle.AddItem "����"
''''            Case 1
''''                cboStyle.AddItem "����"
''''            Case 2
''''                cboStyle.AddItem "����"
''''        End Select
''''        cboStyle.ListIndex = 0
''''    Else
''''        cboStyle.ListIndex = rsTmp!���ӱ�־
''''    End If
''''    txtMoney.Text = Format(rsTmp!ʵ�ս��, "0.00")
''''    txt����Ա.Text = rsTmp!����Ա����
''''    txtDate.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
''''    ReadBill = -1
''''    Exit Function
''''errH:
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''End Function
''''
''''Private Function CancelBill(strNO As String) As Boolean
'''''���ܣ��˳����˾��￨���ü�¼
''''    Dim strSQL As String
''''    Dim blnTrans As Boolean
''''
''''    On Error GoTo errH
''''
''''    '���ù���"zl_���￨��¼_Delete"
''''    strSQL = "zl_���￨��¼_DELETE('" & strNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
''''
''''    gcnOracle.BeginTrans: blnTrans = True
''''
'''''    Call SQLTest(App.ProductName, Me.Caption, strSQL)   'SQLTest
'''''    gcnOracle.Execute strSQL, , adCmdStoredProc
'''''    Call SQLTest
''''    zldatabase.ExecuteProcedure strSQL, Me.Caption
''''
''''    gcnOracle.CommitTrans: blnTrans = False
''''
''''    CancelBill = True
''''    Exit Function
''''errH:
''''    If blnTrans Then gcnOracle.RollbackTrans
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''End Function
''''
''''Private Sub txtPatient_LostFocus()
''''    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
''''End Sub
''''
''''Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    If Button = 2 Then
''''        glngTXTProc = GetWindowLong(txtPatient.hwnd, GWL_WNDPROC)
''''        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, AddressOf WndMessage)
''''    End If
''''End Sub
''''
''''Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    If Button = 2 Then
''''        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, glngTXTProc)
''''    End If
''''End Sub
''''
''''Private Sub txt���￨_GotFocus()
''''    Call zlControl.TxtSelAll(txt���￨)
'''''    With txtApplyman
'''''        .SelStart = 0
'''''        .SelLength = Len(.Text)
'''''    End With
''''End Sub
''''
'''''����28130��27929 by lesfeng 2010-02-26
''''Private Sub txt���￨_KeyPress(KeyAscii As Integer)
''''    Dim strCardNO As String
''''    KeyAscii = Asc(UCase(Chr(KeyAscii)))
''''    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
''''    If KeyAscii = 13 Then
''''        strCardNO = Trim(txt���￨)
''''        If mbytInState = 0 Then
''''            Call NewCard
''''            If ReadCardNo(strCardNO, 2) = -1 Then
''''                cmdOK.SetFocus
''''            Else
''''                Call zlControl.TxtSelAll(txt���￨)
'''''                txt���￨.SetFocus
''''                sta.Panels(2) = "û�з��ָþ��￨����Ϣ,����δ����,���飡"
''''            End If
''''        Else
''''            If ReadCardNo(strCardNO, 1) = -1 Then
''''                cmdOK.SetFocus
''''            Else
''''                Call zlControl.TxtSelAll(txt���￨)
'''''                txt���￨.SetFocus
''''                sta.Panels(2) = "û�з��ָþ��￨����Ϣ,����δ����,���飡"
''''            End If
''''        End If
''''    End If
''''End Sub
''''
''''Private Function ReadCardNo(ByVal strNO As String, ByVal intFlag As Integer) As Integer
'''''���ܣ�ˢ����֤���￨�˿�����һ���Լ�ˢ��ȡ��
'''''���룺strNO ����
'''''      intFlag ��־ 1 ��֤ 2 ȡ��
'''''���أ�
'''''     -1:�ɹ�
'''''      0:ʧ��
'''''      1:�ü�¼������
''''    On Error GoTo errH
''''
''''    Dim rsTmp As New ADODB.Recordset
''''    Dim strSQL As String
''''    Dim lng����ID As Long
''''    Dim str���ݺ� As String
''''
''''    ReadCardNo = 0
''''
''''    strSQL = "select ���￨��,����,����ID from ������Ϣ where ���￨�� = [1]"
''''
''''    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
''''    If rsTmp.EOF Then ReadCardNo = 1: Exit Function
''''
''''    mstr�˿���֤ = IIf(IsNull(rsTmp!���￨��), "", rsTmp!���￨��)
''''    lng����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
''''    If intFlag = 1 Then
''''        ReadCardNo = -1
''''        rsTmp.Close
''''        Exit Function
''''    Else
''''        rsTmp.Close
''''        '��ȡ���￨�ڷ����е�No
''''        strSQL = _
''''        " Select A.NO,B.����֤�� " & _
''''        " From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A,������Ϣ B" & _
''''        " Where A.��¼����=5 And A.����ID=B.����ID And A.ʵ��Ʊ��=[1] and A.����ID = [2]" & _
''''         IIf(mblnViewCancel, " And A.��¼״̬=3 ", " And A.��¼״̬=1 ")
''''         '����31841:" And A.��¼״̬=1 "
''''         Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng����ID)
''''        If rsTmp.EOF Then ReadCardNo = 1: Exit Function
''''
''''        str���ݺ� = IIf(IsNull(rsTmp!NO), "", rsTmp!NO)
''''        If ReadBill(str���ݺ�) = -1 Then
''''            ReadCardNo = -1
''''            rsTmp.Close
''''            Exit Function
''''        End If
''''    End If
''''    rsTmp.Close
''''    Exit Function
''''errH:
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''End Function
''''
''''Private Sub txt���￨_LostFocus()
''''    Dim strCardNO As String
''''
''''    strCardNO = Trim(txt���￨)
''''    If mbytInState = 0 Then
''''        Call NewCard
''''        If ReadCardNo(strCardNO, 2) = -1 Then
''''            cmdOK.SetFocus
''''        Else
'''''            txt���￨.SetFocus
''''            sta.Panels(2) = "û�з��ָþ��￨����Ϣ,����δ����,���飡"
''''        End If
''''    Else
''''        If ReadCardNo(strCardNO, 1) = -1 Then
''''            cmdOK.SetFocus
''''        Else
'''''            txt���￨.SetFocus
''''            sta.Panels(2) = "û�з��ָþ��￨����Ϣ,����δ����,���飡"
''''        End If
''''    End If
''''End Sub
