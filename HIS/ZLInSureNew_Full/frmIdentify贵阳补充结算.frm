VERSION 5.00
Begin VB.Form frmIdentify����������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����������"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frmIdentify�����������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&E)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5445
      TabIndex        =   26
      Top             =   4335
      Width           =   1335
   End
   Begin VB.Frame frmDetail 
      Height          =   4035
      Left            =   195
      TabIndex        =   30
      Top             =   150
      Width           =   8070
      Begin VB.TextBox txtҽ���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1515
         TabIndex        =   3
         Top             =   960
         Width           =   2160
      End
      Begin VB.TextBox txt��ҳID 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1485
         Width           =   2445
      End
      Begin VB.TextBox txt����ID 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   2445
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1965
         Width           =   2445
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1905
         Width           =   2445
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5385
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   960
         Width           =   2160
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "����"
         Height          =   375
         Left            =   1515
         TabIndex        =   0
         Top             =   285
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton optסԺ 
         Caption         =   "סԺ"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5385
         TabIndex        =   1
         Top             =   285
         Width           =   810
      End
      Begin VB.CommandButton cmd���� 
         Height          =   300
         Left            =   7530
         Picture         =   "frmIdentify�����������.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   300
      End
      Begin VB.CommandButton cmdҽ���� 
         Height          =   300
         Left            =   3660
         Picture         =   "frmIdentify�����������.frx":00EA
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   300
      End
      Begin VB.TextBox txt����˳��� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1515
         TabIndex        =   21
         Top             =   3060
         Width           =   2445
      End
      Begin VB.TextBox txt������ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5385
         TabIndex        =   23
         Top             =   3075
         Width           =   2445
      End
      Begin VB.ComboBox cbo֧����� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1515
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3540
         Width           =   2445
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2430
         Width           =   2445
      End
      Begin VB.TextBox txt���֤�� 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2430
         Width           =   2445
      End
      Begin VB.Label labҽ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����(&Y)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   495
         TabIndex        =   2
         Top             =   1005
         Width           =   945
      End
      Begin VB.Label lab��ҳID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҳID(&P)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4320
         TabIndex        =   10
         Top             =   1530
         Width           =   945
      End
      Begin VB.Label lab����ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ID(&S)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   495
         TabIndex        =   8
         Top             =   1485
         Width           =   945
      End
      Begin VB.Label lab���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&I)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4530
         TabIndex        =   5
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label labסԺ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��(&H)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   210
         Left            =   4320
         TabIndex        =   14
         Top             =   2010
         Width           =   945
      End
      Begin VB.Label lab���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   705
         TabIndex        =   12
         Top             =   1950
         Width           =   735
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   60000
         Y1              =   2895
         Y2              =   2895
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0080FFFF&
         X1              =   0
         X2              =   60000
         Y1              =   2910
         Y2              =   2910
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0080FFFF&
         X1              =   0
         X2              =   60000
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   60000
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lab����˳��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����˳���(&J)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   75
         TabIndex        =   20
         Top             =   3105
         Width           =   1365
      End
      Begin VB.Label lab������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������(&B)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4125
         TabIndex        =   22
         Top             =   3120
         Width           =   1155
      End
      Begin VB.Label lbl֧����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "֧�����(&T)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   285
         TabIndex        =   24
         Top             =   3600
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�(&X)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   705
         TabIndex        =   16
         Top             =   2475
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��(&F)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   210
         Left            =   4110
         TabIndex        =   18
         Top             =   2475
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6930
      TabIndex        =   27
      Top             =   4335
      Width           =   1335
   End
   Begin VB.PictureBox P2 
      Height          =   495
      Left            =   1530
      Picture         =   "frmIdentify�����������.frx":014A
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   29
      Top             =   7035
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox P1 
      Height          =   495
      Left            =   165
      Picture         =   "frmIdentify�����������.frx":0228
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   28
      Top             =   7035
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmIdentify�����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure      As Integer

Private Sub cmdCancel_Click()
    With g�������
        .blnYn = False
    End With
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrH
    Dim rsTmp                   As ADODB.Recordset
    Dim str����˳���           As String
    Dim str������             As String
    Dim str֧������             As String
    Dim lng����ID               As String
    Dim strMsg                  As String
    
    str����˳��� = txt����˳���.Text
    str������ = txt������.Text
    lng����ID = Val(txt����ID.Text)
    str֧������ = cbo֧�����.ItemData(cbo֧�����.ListIndex)
    '��������ڱ����Ƿ���� �����������ɾ��
    gstrSQL = "Select count(*) as Cnt From ���ս����¼ where ֧��˳���=[1] and ��ע=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str������, IIf(cbo֧�����.Text = "��ͨ����", "��ͨ", "����") & str����˳���)
    If rsTmp!cnt > 0 Then
        strMsg = "����Ҫ�����ĵ�����Ϣ" & vbCrLf
        strMsg = strMsg & "����˳��ţ���" & str����˳��� & "��" & vbCrLf
        strMsg = strMsg & "�����ţ���" & txt������ & "��" & vbCrLf
        strMsg = strMsg & "֧�����ͣ���" & cbo֧�����.Text & "��" & vbCrLf
        strMsg = strMsg & "��HISϵͳ���Ѵ��ڣ��뵽HIS�շ�ϵͳ�н��г�����"
        MsgBox strMsg, vbCritical, gstrSysName
        Exit Sub
    End If
    strMsg = "�Ƿ�������ݣ���Ҫ�����ĵ�����Ϣ���£�" & vbCrLf
    strMsg = strMsg & "����˳��ţ���" & str����˳��� & "��" & vbCrLf
    strMsg = strMsg & "�����ţ���" & txt������ & "��" & vbCrLf
    strMsg = strMsg & "֧�����ͣ���" & cbo֧�����.Text & "��" & vbCrLf
    strMsg = strMsg & "ѡ��[��]���������ĵĵ��ݣ�"
    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "BILLNO", str����˳���)     ' ����˳���
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str������)    ' ������
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str֧������)     ' ֧�����
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.����)     ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))  ' ��������
    
    '���ýӿ�
    If MsgBox("���ٴ�ȷ���Ƿ���Ʊ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If CommServer("RETBALANCE", IIf(IS����(lng����ID), 1, 0)) Then
        '���������Ϣ
        gstrSQL = "ZL_���������շ�_Update('" & str����˳��� & "','" & str������ & "','" & str֧������ & "��" & cbo֧�����.Text & "','" & txtҽ����.Text & "','" & txt����.Text & "','" & txt����ID.Text & "','" & IIf(opt����.Value, 0, txt��ҳID.Text) & "','" & txt����.Text & "','" & txt�Ա�.Text & "','" & txt���֤��.Text & "','" & IIf(opt����.Value, 0, txtסԺ��.Text) & "','" & UserInfo.���� & "','" & zlDatabase.Currentdate & "' ,'�ɹ�')"
        MsgBox "�����������ݳɹ���", vbExclamation, gstrSysName
    Else
        gstrSQL = "ZL_���������շ�_Update('" & str����˳��� & "','" & str������ & "','" & str֧������ & "��" & cbo֧�����.Text & "','" & txtҽ����.Text & "','" & txt����.Text & "','" & txt����ID.Text & "','" & IIf(opt����.Value, 0, txt��ҳID.Text) & "','" & txt����.Text & "','" & txt�Ա�.Text & "','" & txt���֤��.Text & "','" & IIf(opt����.Value, 0, txtסԺ��.Text) & "','" & UserInfo.���� & "','" & zlDatabase.Currentdate & "' ,'ʧ��')"
        MsgBox "������������ʧ�ܣ�", vbExclamation, gstrSysName
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���������շ�")
    Exit Sub
ErrH:
    MsgBox "������������ʧ�ܣ�" & vbCrLf & Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

Private Sub cmd����_Click()
    
    Dim rsTmp       As ADODB.Recordset
    
    cmd����.Picture = P2.Picture
    txt����.Locked = False
    txt����.ForeColor = vbBlue
    txt����.BackColor = &HC0FFC0
    txt����.SetFocus
    txt����.SelStart = 0
    txt����.SelLength = Len(txt����.Text)
    cmdҽ����.Picture = P1.Picture
    txtҽ����.Locked = True
    txtҽ����.ForeColor = &HFF00FF
    txtҽ����.BackColor = vbWhite
    
    If Trim(txt����.Text) = "" Then Exit Sub
    
    gstrSQL = "select * from ������Ϣ where ����=[1] And IC����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintInsure, txt����.Text)
    If rsTmp.RecordCount = 1 Then
        txt����.Text = "" & rsTmp!IC����
        txtҽ����.Text = "" & rsTmp!ҽ����
        txt����ID.Text = "" & rsTmp!����ID
        txt��ҳID = 0
        txt����.Text = rsTmp!����
        txtסԺ��.Text = "" & rsTmp!סԺ��
        txt�Ա�.Text = "" & rsTmp!�Ա�
        txt���֤��.Text = "" & rsTmp!���֤��
        cmdOK.Enabled = True
    Else
        MsgBox "ĩ�ҵ���ز�����Ϣ", vbCritical, gstrSysName
        cmdOK.Enabled = False
    End If
    
End Sub

Private Sub cmdҽ����_Click()
    Dim rsTmp       As ADODB.Recordset
    
    cmdҽ����.Picture = P2.Picture
    txtҽ����.Locked = False
    txtҽ����.ForeColor = vbBlue
    txtҽ����.BackColor = &HC0FFC0
    txtҽ����.SetFocus
    txtҽ����.SelStart = 0
    txtҽ����.SelLength = Len(txtҽ����.Text)
    cmd����.Picture = P1.Picture
    txt����.Locked = True
    txt����.ForeColor = &HFF00FF
    txt����.BackColor = vbWhite
    
    If Trim(txtҽ����) = "" Then Exit Sub
    
    gstrSQL = "select * from ������Ϣ where ����=[1] And ҽ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintInsure, txtҽ����.Text)
    If rsTmp.RecordCount = 1 Then
        txt����.Text = "" & rsTmp!IC����
        txtҽ����.Text = "" & rsTmp!ҽ����
        txt����ID.Text = "" & rsTmp!����ID
        txt��ҳID = 0
        txt����.Text = rsTmp!����
        txtסԺ��.Text = "" & rsTmp!סԺ��
        txt�Ա�.Text = "" & rsTmp!�Ա�
        txt���֤��.Text = "" & rsTmp!���֤��
        cmdOK.Enabled = True
    Else
        MsgBox "ĩ�ҵ���ز�����Ϣ", vbCritical, gstrSysName
        cmdOK.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    Call cmdҽ����_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    lab��ҳID.ForeColor = &H80000003
    txt��ҳID.Enabled = False
    labסԺ��.ForeColor = &H80000003
    txtסԺ��.Enabled = False
    With cbo֧�����
        .AddItem "��ͨ����"
        .ItemData(.NewIndex) = 11
        .AddItem "��������"
        .ItemData(.NewIndex) = 18
        .ListIndex = 0
    End With
End Sub

Private Sub opt����_Click()
    Call ControlRefech
End Sub

Private Sub optסԺ_Click()
    Call ControlRefech
End Sub

Private Sub txt����ID_Validate(Cancel As Boolean)
    If Val(txt����ID.Text) <= 0 Then Cancel = True Else txt����ID.Text = Val(txt����ID.Text)
End Sub

Private Sub txt��ҳID_Validate(Cancel As Boolean)
    If Val(txt��ҳID.Text) <= 0 Then Cancel = True Else txt��ҳID.Text = Val(txt��ҳID.Text)
End Sub

Public Sub ControlRefech()
On Error GoTo ErrH
    If opt����.Value Then
        lab��ҳID.ForeColor = &H80000003
        txt��ҳID.Enabled = False
        labסԺ��.ForeColor = &H80000003
        txtסԺ��.Enabled = False
        
        With cbo֧�����
            .Clear
            .AddItem "��ͨ����"
            .ItemData(.NewIndex) = 11
            .AddItem "��������"
            .ItemData(.NewIndex) = 18
            .ListIndex = 0
        End With
    Else
        lab��ҳID.ForeColor = vbBlack
        txt��ҳID.Enabled = True
        labסԺ��.ForeColor = vbBack
        txtסԺ��.Enabled = True
        With cbo֧�����
            .Clear
            .AddItem "��ͨסԺ"
            .ItemData(.NewIndex) = 31
            .ListIndex = 0
        End With
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub
 

Public Property Let Insure(ByVal vNewValue As Integer)
    mintInsure = vNewValue
End Property

