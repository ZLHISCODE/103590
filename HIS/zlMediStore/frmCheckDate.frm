VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheckDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�̵���������"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmCheckDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   3120
      TabIndex        =   16
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   300
      Left            =   1950
      TabIndex        =   0
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   300
      Left            =   360
      TabIndex        =   17
      Top             =   1440
      Width           =   1100
   End
   Begin VB.Frame fraCondition 
      Caption         =   "����"
      Height          =   1200
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox pic���龫�� 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   960
         ScaleHeight     =   615
         ScaleWidth      =   3015
         TabIndex        =   11
         Top             =   1800
         Width           =   3015
         Begin VB.CheckBox chkҩƷ���� 
            Caption         =   "����II��"
            Height          =   180
            Index           =   3
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkҩƷ���� 
            Caption         =   "����ҩ"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkҩƷ���� 
            Caption         =   "����I��"
            Height          =   180
            Index           =   1
            Left            =   1440
            TabIndex        =   13
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox chkҩƷ���� 
            Caption         =   "����ҩ"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.PictureBox pic��Ч�� 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   960
         ScaleHeight     =   495
         ScaleWidth      =   3015
         TabIndex        =   6
         Top             =   1080
         Width           =   3015
         Begin VB.TextBox txt 
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1440
            TabIndex        =   9
            Text            =   "30"
            Top             =   270
            Width           =   300
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "    ��"
            Height          =   255
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkʧЧ 
            Caption         =   "��ʧЧ"
            Height          =   255
            Left            =   2160
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblЧ�� 
            AutoSize        =   -1  'True
            Caption         =   "��Ч��ʱ��"
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   720
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   185794563
         CurrentDate     =   36901
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "�̵�����"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�̵�ʱ��"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCheckDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mint�̵�ʱ�䷶Χ As Integer
Private mstr�̵�ʱ�� As String
Private mint�༭״̬ As Integer             '7���ⷿȫ��ҩƷ�̵㣻8������ҩƷ�̵㣻9���Զ���������������δ�̵��ҩƷ
Private mint�̵����� As Integer
Private mstr��Ч�� As String
Private mstr���� As String

Public Function GetCondition(FrmMain As Form, ByRef str�̵�ʱ�� As String, ByVal int�༭״̬ As Integer, ByRef int�̵����� As Integer, ByRef str��Ч�� As String, ByRef str���� As String) As Boolean
    
    mblnReturn = False
    mint�༭״̬ = int�༭״̬
    
    Me.Show 1, FrmMain
    
    str�̵�ʱ�� = mstr�̵�ʱ��
    int�̵����� = mint�̵�����
    str��Ч�� = mstr��Ч��
    str���� = mstr����
    
    GetCondition = mblnReturn
    
End Function


Private Sub cboType_Click()
    'ѡ���Ч�ڿɼ�
    pic��Ч��.Visible = cboType.ListIndex = 0
    If pic��Ч��.Visible Then pic��Ч��.Top = 1080
    'ѡ���龫��ɼ�
    pic���龫��.Visible = cboType.ListIndex = 1
    If pic���龫��.Visible Then pic���龫��.Top = 1080
    
    If cboType.ListIndex > 1 Then '��������
        fraCondition.Height = 1200
        cmdHelp.Top = fraCondition.Top + fraCondition.Height + 120
        CmdSave.Top = cmdHelp.Top
        CmdCancel.Top = cmdHelp.Top
        Me.Height = cmdHelp.Top + cmdHelp.Height + 520 '�ı䴰��߶�
    Else
        fraCondition.Height = 1800
        cmdHelp.Top = fraCondition.Top + fraCondition.Height + 120
        CmdSave.Top = cmdHelp.Top
        CmdCancel.Top = cmdHelp.Top
        Me.Height = cmdHelp.Top + cmdHelp.Height + 520 '�ı䴰��߶�
    End If
    
End Sub


Private Sub CmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub CmdSave_Click()
    mblnReturn = True
    mstr�̵�ʱ�� = Format(dtpDate.Value, "yyyy-MM-dd hh:mm:ss")
    If cboType.ListIndex = 0 And chkDay.Value = 0 And chkʧЧ.Value = 0 Then
        MsgBox "��Խ�Ч��ʱ��������ã�", vbInformation + vbOKOnly, gstrSysName
        chkDay.SetFocus
        Exit Sub
    ElseIf cboType.ListIndex = 1 And (chkҩƷ����(0).Value = 0 And chkҩƷ����(1).Value = 0 And chkҩƷ����(2).Value = 0 And chkҩƷ����(3).Value = 0) Then
        MsgBox "��ѡ���̵�ҩƷ�������ͣ�", vbInformation + vbOKOnly, gstrSysName
        chkҩƷ����(0).SetFocus
        Exit Sub
    End If
    
    mint�̵����� = cboType.ListIndex
    mstr��Ч�� = IIf(chkDay.Value = 0, 0, Val(txt.Text)) & ":" & chkʧЧ.Value '��Ч�ڷ�������
    mstr���� = chkҩƷ����(0).Value & ":" & chkҩƷ����(2).Value & ":" & chkҩƷ����(1).Value & ":" & chkҩƷ����(3).Value '����˳�����������ԡ�����I�ࡢ����II��
    
    Unload Me
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    mint�̵�ʱ�䷶Χ = Val(zlDataBase.GetPara("�̵�ʱ�䷶Χ����", glngSys, 1307, 30))
    dtpDate.MinDate = CDate(Format(DateAdd("d", -mint�̵�ʱ�䷶Χ, Date), "yyyy-mm-dd") & " 00:00:00")
    'ҩƷ����Ȩ�޿���
    
    dtpDate.Value = Format(Sys.Currentdate, dtpDate.CustomFormat)
    dtpDate.MaxDate = dtpDate.Value
    
    If mint�༭״̬ = 8 Then '����ҩƷ�̵�
        cboType.AddItem "0-��Ч��ҩƷ"
        cboType.AddItem "1-���龫��ҩƷ"
        cboType.AddItem "2-ͣ��ҩƷ"
'        cboType.AddItem "3-�޿���¼��ҩƷ"
        cboType.AddItem "3-���������п������۵�ҩƷ"
        cboType.AddItem "4-����ҩ��"
        
        cboType.ListIndex = 0
    ElseIf mint�༭״̬ = 7 Or mint�༭״̬ = 9 Then
        fraCondition.Height = 720
        cmdHelp.Top = fraCondition.Top + fraCondition.Height + 120
        CmdSave.Top = cmdHelp.Top
        CmdCancel.Top = cmdHelp.Top
        Me.Height = cmdHelp.Top + cmdHelp.Height + 520 '�ı䴰��߶�
    End If
End Sub


Private Sub txt_Change()
    If (Val(txt.Text) <= 0 Or Val(txt.Text) > 999) And txt.Text <> "" Then
        txt.Text = 30
    End If
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
