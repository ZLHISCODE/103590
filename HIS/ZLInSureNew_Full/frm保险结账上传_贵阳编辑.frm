VERSION 5.00
Begin VB.Form frm���ս����ϴ�_�����༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ҽ�����������ϴ�"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5983.333
   ScaleMode       =   0  'User
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fra���������༭ 
      Caption         =   "���������ϴ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   323
      TabIndex        =   2
      Top             =   210
      Width           =   5715
      Begin VB.TextBox Txt��¼id 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   25
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txt�շ�ʱ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3660
         Width           =   2775
      End
      Begin VB.TextBox txt�������� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3285
         Width           =   2775
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1740
         Width           =   2775
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txt�Ա� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2130
         Width           =   2775
      End
      Begin VB.TextBox txtסԺ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txt����˳��� 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1350
         Width           =   2775
      End
      Begin VB.TextBox txt��ʼ���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2910
         Width           =   2775
      End
      Begin VB.TextBox txt֧����� 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   975
         Width           =   2775
      End
      Begin VB.CommandButton Cmd����ID 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4980
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   285
      End
      Begin VB.TextBox txtҽ���� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2505
         MaxLength       =   14
         TabIndex        =   8
         Top             =   232
         Width           =   2775
      End
      Begin VB.Label Lab��¼id 
         Caption         =   "��¼id"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lab�շ�ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "�շ�ʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   23
         Top             =   3720
         Width           =   780
      End
      Begin VB.Label lab�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   22
         Top             =   3345
         Width           =   780
      End
      Begin VB.Label lab���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1605
         TabIndex        =   19
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label lab�Ա� 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1605
         TabIndex        =   18
         Top             =   2190
         Width           =   390
      End
      Begin VB.Label lab���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1605
         TabIndex        =   17
         Top             =   2580
         Width           =   390
      End
      Begin VB.Label labסԺ�� 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   13
         Top             =   660
         Width           =   585
      End
      Begin VB.Label lab��ʼ���� 
         AutoSize        =   -1  'True
         Caption         =   "��ʼ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   12
         Top             =   2970
         Width           =   780
      End
      Begin VB.Label lab����˳��� 
         AutoSize        =   -1  'True
         Caption         =   "����˳���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1020
         TabIndex        =   11
         Top             =   1410
         Width           =   975
      End
      Begin VB.Label lab֧����� 
         AutoSize        =   -1  'True
         Caption         =   "֧�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   10
         Top             =   1035
         Width           =   780
      End
      Begin VB.Label labҽ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����(&B)*"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1125
         TabIndex        =   9
         Top             =   285
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4845
      TabIndex        =   1
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "���沢�ϴ�(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2805
      TabIndex        =   0
      Top             =   4785
      Width           =   1500
   End
   Begin VB.Shape sapStatus 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frm���ս����ϴ�_�����༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr����סԺ            As String
Private mintInsure              As Integer
Private mblnOkCancel            As Boolean
Private mstr����˳���          As String
Dim sngX                        As Single
Dim sngY                        As Single
Dim sngH                        As Single
Dim strUp                       As String

Public Property Let Insure(ByVal vNewValue As Integer)
    mintInsure = vNewValue
End Property

Public Property Get OkCancel() As Boolean
    OkCancel = mblnOkCancel
End Property

Public Property Get ����˳���() As String
    ����˳��� = mstr����˳���
End Property

Public Property Let ����˳���(ByVal vNewValue As String)
    mstr����˳��� = vNewValue
End Property

Private Sub cmdADD_Click()
    Dim strDateTime As String
On Error GoTo ErrH
    mstr����˳��� = txt����˳���.Text
    If Not mCheckData Then Exit Sub
     
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", txt֧�����.Text)    ' ֧�����
    Call InsertChild(mdomInput.documentElement, "BILLNO", mstr����˳���)       ' ����˳���
    '���ýӿ�
    If CommServer("UPLOADBYBILLNO") = False Then Exit Sub

    '����
    gstrSQL = "Zl_����_�����ϴ�_Update('" & txt֧�����.Text & "','" & mstr����˳��� & "','" & txtҽ����.Tag & "','" & txtסԺ��.Text & "','" & txt����.Text & "','" & txt�Ա�.Text & "','" & txt����.Text & "'," & IIf(IsDate(txt��ʼ����.Text), "to_date('" & txt��ʼ����.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & "," & IIf(IsDate(txt��������.Text), "to_date('" & txt��������.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & "," & IIf(IsDate(txt�շ�ʱ��.Text), "to_date('" & txt�շ�ʱ��.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & ",'" & UserInfo.���� & "',sysdate,'" & Txt��¼id.Text & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
'
    mblnOkCancel = True
    txtҽ����.Text = ""
    txtҽ����.Tag = ""
    txt����.Text = ""
    sapStatus.Visible = True
    sapStatus.FillColor = &HC000&
    sapStatus.BorderColor = &HC000&
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    sapStatus.FillColor = vbRed
    sapStatus.BorderColor = vbRed
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrH
    Unload Me
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dim strDateTime     As String
On Error GoTo ErrH
    mstr����˳��� = txt����˳���.Text
    If Not mCheckData Then Exit Sub
     
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", txt֧�����.Text)    ' ֧�����
    Call InsertChild(mdomInput.documentElement, "BILLNO", mstr����˳���)       ' ����˳���
    '���ýӿ�
    If CommServer("UPLOADBYBILLNO") = False Then Exit Sub

    '����
    gstrSQL = "Zl_����_�����ϴ�_Update('" & txt֧�����.Text & "','" & mstr����˳��� & "','" & txtҽ����.Tag & "','" & txtסԺ��.Text & "','" & txt����.Text & "','" & txt�Ա�.Text & "','" & txt����.Text & "'," & IIf(IsDate(txt��ʼ����.Text), "to_date('" & txt��ʼ����.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & "," & IIf(IsDate(txt��������.Text), "to_date('" & txt��������.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & "," & IIf(IsDate(txt�շ�ʱ��.Text), "to_date('" & txt�շ�ʱ��.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & ",'" & UserInfo.���� & "',sysdate,'" & Txt��¼id.Text & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
'
    mblnOkCancel = True
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Err.Clear
    Exit Sub
End Sub

Private Sub Cmd����ID_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim vRect   As RECT
    
    'gstrSQL = "select /*+ rule */C.������ˮ�� as ID,d.id as ��¼id,B.ҽ����,B.סԺ��,C.ҽ����� As ֧�����,C.������ˮ�� as ����˳���,B.����,B.�Ա�,B.����, A.��Ժ����,A.��Ժ����,D.�շ�ʱ��" & vbNewLine & _
       '         "from ������ҳ A,������Ϣ B,���ս����¼ C,���˽��ʼ�¼ D" & vbNewLine & _
        '        "where A.��Ժ����>sysdate-[2] And C.����=2 And A.����ID=B.����ID And A.����ID=C.����ID And C.��¼id = D.Id And B.סԺ�� is not null And A.����=[1]" & vbNewLine & _
        '   '     "And C.������ˮ�� Not In (Select ����˳��� From ����_�����ϴ�)"
    
    gstrSQL = "Select /*+ rule */ Distinct c.������ˮ�� As Id, a.Id As ��¼id, b.ҽ����, b.סԺ��,b.���� , b.�Ա�, b.����, a.�շ�ʱ��, c.֧����� As ֧�����, c.������ˮ�� As ����˳���," & vbNewLine & _
            "  a.��ʼ����, a.��������" & vbNewLine & _
            "From ���˽��ʼ�¼ a, ������Ϣ b," & vbNewLine & _
           "(Select Distinct b.סԺ��, b.����id, a.������ˮ��, a.��¼id, a.ҽ����� As ֧�����, a.����ʱ�� " & vbNewLine & _
            "From ���ս����¼ a, ������ҳ b" & vbNewLine & _
            "Where a. ����id = b.����id And a.����ʱ�� >=sysdate-1000  And b.���� = [1] And a.���� = 2) c" & vbNewLine & _
            "Where a.����id = c.����id And a.Id = c.��¼id And a.�շ�ʱ�� >= Sysdate - 1000 And c.����id = b.����id And b.���� = [1] And " & vbNewLine & _
           " a.��¼״̬ = 1 And a.�շ�ʱ�� >= Sysdate -1000  And C.������ˮ�� Not In (Select ����˳��� From ����_�����ϴ�)  " & vbNewLine & _
           " order by b.ҽ����"
    
    
    vRect = GetControlRect(txtҽ����.hwnd)
    sngX = vRect.Left
    sngY = vRect.Top
    sngH = txtҽ����.Height
 
    DoEvents
    Set rsTemp = zlDatabase.ShowSQLSelect( _
            Nothing, gstrSQL, 0, "������Ϣ��ѯ", False, _
            "", "", False, False, True, _
            sngX, sngY, sngH, False, False, _
            False, mintInsure, 90 _
            )
    If ChkRsState(rsTemp) Then
        txtҽ����.Tag = ""
        txtҽ����.Text = ""
        txtסԺ��.Text = ""
        txt����.Text = ""
        txt�Ա�.Text = ""
        txt����.Text = ""
        txt֧�����.Text = ""
        txt����˳���.Text = ""
        Txt��¼id.Text = ""
        txt��ʼ����.Text = ""
        txt��������.Text = ""
        txt�շ�ʱ��.Text = ""
    Else
        txtҽ����.Tag = Nvl(rsTemp!ҽ����)
        txtҽ����.Text = Nvl(rsTemp!ҽ����)
        txtסԺ��.Text = Nvl(rsTemp!סԺ��)
        txt����.Text = Nvl(rsTemp!����)
        txt�Ա�.Text = Nvl(rsTemp!�Ա�)
        txt����.Text = Nvl(rsTemp!����)
        txt֧�����.Text = Nvl(rsTemp!֧�����)
        txt����˳���.Text = Nvl(rsTemp!����˳���)
        Txt��¼id.Text = Nvl(rsTemp!��¼ID)
        txt��ʼ����.Text = Format(Nvl(rsTemp!��ʼ����), "yyyy-mm-dd hh:mm:ss")
        txt��������.Text = Format(Nvl(rsTemp!��������), "yyyy-mm-dd hh:mm:ss")
        txt�շ�ʱ��.Text = Format(Nvl(rsTemp!�շ�ʱ��), "yyyy-mm-dd hh:mm:ss")
        
        zlCommFun.PressKey vbKeyTab
End If
End Sub

Private Sub txt�Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtҽ����_KeyPress(KeyAscii As Integer)
    Dim rsTemp      As New ADODB.Recordset
    Dim vRect       As RECT
    Dim strText     As String
    
    If KeyAscii <> 13 Then Exit Sub
    If Trim(txtҽ����.Text) = "" Then Exit Sub
    If txtҽ����.Locked Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
    strText = txtҽ����.Text
  '  gstrSQL = "select /*+ rule */ C.������ˮ�� as ID,d.id as ��¼id,B.ҽ����,B.סԺ��,C.ҽ����� As ֧�����,C.������ˮ�� as ����˳���,B.����,B.�Ա�,B.����, A.��Ժ����,A.��Ժ����,D.�շ�ʱ��" & vbNewLine & _
   '             "from ������ҳ A,������Ϣ B,���ս����¼ C,���˽��ʼ�¼ D" & vbNewLine & _
     '           "where A.��Ժ����>sysdate-[2] And C.����=2 And A.����ID=B.����ID And A.����ID=C.����ID And C.��¼id = D.Id And B.סԺ�� is not null And A.����=[1]" & vbNewLine & _
      '          "And C.������ˮ�� Not In (Select ����˳��� From ����_�����ϴ�)"
                
                
     gstrSQL = "Select /*+ rule */ distinct c.������ˮ�� As Id, a.Id As ��¼id, b.ҽ����, b.סԺ��,b.���� , b.�Ա�, b.����, a.�շ�ʱ��, c.֧����� As ֧�����, c.������ˮ�� As ����˳���," & vbNewLine & _
            "  a.��ʼ����, a.��������" & vbNewLine & _
            "From ���˽��ʼ�¼ a, ������Ϣ b," & vbNewLine & _
           "(Select Distinct b.סԺ��, b.����id, a.������ˮ��, a.��¼id, a.ҽ����� As ֧�����, a.����ʱ�� " & vbNewLine & _
            "From ���ս����¼ a, ������ҳ b" & vbNewLine & _
            "Where a. ����id = b.����id And a.����ʱ�� >=sysdate-1000  And b.���� = [1] And a.���� = 2) c" & vbNewLine & _
            "Where a.����id = c.����id And a.Id = c.��¼id And a.�շ�ʱ�� >= Sysdate - 1000 And c.����id = b.����id And b.���� = [1] And " & vbNewLine & _
           " a.��¼״̬ = 1 And a.�շ�ʱ�� >= Sysdate -1000  And C.������ˮ�� Not In (Select ����˳��� From ����_�����ϴ�)"
           
    If zlCommFun.IsCharAlpha(strText) Then
        gstrSQL = gstrSQL & vbCrLf & "And zlspellcode(B.����) like '" & UCase(strText) & "%'"
    ElseIf zlCommFun.IsNumOrChar(strText) Then
        gstrSQL = gstrSQL & vbCrLf & "And (B.סԺ�� like '" & UCase(strText) & "%' or B.ҽ���� like  '" & UCase(strText) & "%')"
    ElseIf zlCommFun.IsCharChinese(strText) Then
        gstrSQL = gstrSQL & vbCrLf & "And B.���� like '" & UCase(strText) & "%'"
    Else
        gstrSQL = gstrSQL & vbCrLf & "And B.ҽ���� like '" & UCase(strText) & "%'"
    End If
    
    vRect = GetControlRect(txtҽ����.hwnd)
    sngX = vRect.Left
    sngY = vRect.Top
    sngH = txtҽ����.Height
    
    'DoEvents
    Set rsTemp = zlDatabase.ShowSQLSelect( _
            Nothing, gstrSQL, 0, "������Ϣ��ѯ", False, _
            "", "", False, False, True, _
            sngX, sngY, sngH, False, False, _
            False, mintInsure, 90 _
            )
    If ChkRsState(rsTemp) Then
        txtҽ����.Tag = ""
        txtҽ����.Text = ""
        txtסԺ��.Text = ""
        txt����.Text = ""
        txt�Ա�.Text = ""
        txt����.Text = ""
        txt֧�����.Text = ""
        txt����˳���.Text = ""
        Txt��¼id.Text = ""
        txt��ʼ����.Text = ""
        txt��������.Text = ""
        txt�շ�ʱ��.Text = ""
    Else
        txtҽ����.Tag = Nvl(rsTemp!ҽ����)
        txtҽ����.Text = Nvl(rsTemp!ҽ����)
        txtסԺ��.Text = Nvl(rsTemp!סԺ��)
        txt����.Text = Nvl(rsTemp!����)
        txt�Ա�.Text = Nvl(rsTemp!�Ա�)
        txt����.Text = Nvl(rsTemp!����)
        txt֧�����.Text = Nvl(rsTemp!֧�����)
        txt����˳���.Text = Nvl(rsTemp!����˳���)
        Txt��¼id.Text = Nvl(rsTemp!��¼ID)
        txt��ʼ����.Text = Format(Nvl(rsTemp!��ʼ����), "yyyy-mm-dd hh:mm:ss")
        txt��������.Text = Format(Nvl(rsTemp!��������), "yyyy-mm-dd hh:mm:ss")
        txt�շ�ʱ��.Text = Format(Nvl(rsTemp!�շ�ʱ��), "yyyy-mm-dd hh:mm:ss")
        zlCommFun.PressKey vbKeyTab
End If
End Sub

Private Sub Form_Load()
    Dim rsTmp       As ADODB.Recordset
    Dim strDate     As String
    
    strDate = zlDatabase.Currentdate
    If mstr����˳��� <> "" Then

'        gstrSQL = "Select /*+ rule */" & vbCrLf & _
'                "to_char(A.����) || to_char(A.����ID) || to_char(A.ҽ����) AS ID," & vbCrLf & _
'                "A.����,A.����ID,B.���� AS ��������,C.���� AS ���ֱ���,C.���� as ��������," & vbCrLf & _
'                "A.ҽ����,A.����,A.�Ա�,A.����," & vbCrLf & _
'                "A.��ע, A.�Ǽ���,A.�Ǽ�����," & vbCrLf & _
'                "A.ȡ���� , A.ȡ������, A.ȡ��ԭ��" & vbCrLf & _
'                "FROM ����_�ز���Ա A,������� B,���ղ��� C" & vbCrLf & _
'                "Where A.���� = B.��� And A.���� = C.���� And A.����ID = C.ID" & vbCrLf & _
'                "And A.����=[1] and A.����ID=[2] And A.ҽ���� =[3]"
'        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintInsure, mstr����˳���, mstr����˳���)
'        If Not ChkRsState(rsTmp) Then
'            '�޸�״̬
'            cmdAdd.Visible = False
'            Cmd����ID.Enabled = False
'        End If
    End If
    txt����.BackColor = gconLockColor
    txt�Ա�.BackColor = gconLockColor
End Sub

Private Function mCheckData() As Boolean
On Error GoTo ErrH
    
    If txtҽ����.Tag = "" Then
        MsgBox "����˳��Ų���Ϊ�գ�", vbCritical, gstrSysName
        Exit Function
    End If
    '��⵱ǰ���Ժ��Ƿ�������ݣ�������������¼�
    gstrSQL = "Select count(1) From ����_�����ϴ� where ����˳���=[1]"
    If zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr����˳���).Fields(0) > 0 Then
        MsgBox "����˳��š�" & mstr����˳��� & "���ѵǼǴ��ڣ�" & vbCrLf & "����������¼���༭��", vbCritical, gstrSysName
        Exit Function
    End If
    
    If MsgBox("��ȷ������¼������ݣ��ϴ��󽫲����޸ģ�", vbOKCancel + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbOK Then
        Exit Function
    End If
    mCheckData = True
    Exit Function
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Function
End Function




