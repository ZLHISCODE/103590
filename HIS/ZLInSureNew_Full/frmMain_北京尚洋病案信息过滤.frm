VERSION 5.00
Begin VB.Form frmMain_�������󲡰���Ϣ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������󲡰���Ϣ����"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5100
   Icon            =   "frmMain_�������󲡰���Ϣ����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   2055
      Left            =   240
      TabIndex        =   2
      Top             =   225
      Width           =   4590
      Begin VB.OptionButton opt3 
         Caption         =   "����"
         Height          =   255
         Left            =   3450
         TabIndex        =   10
         Top             =   1440
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton opt2 
         Caption         =   "δ�ϴ�"
         Height          =   225
         Left            =   2490
         TabIndex        =   9
         Top             =   1455
         Width           =   885
      End
      Begin VB.OptionButton opt1 
         Caption         =   "���ϴ�"
         Height          =   195
         Left            =   1605
         TabIndex        =   8
         Top             =   1455
         Width           =   930
      End
      Begin VB.TextBox txt��ˮ�� 
         Appearance      =   0  'Flat
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
         Left            =   1590
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtҽ���� 
         Appearance      =   0  'Flat
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
         Left            =   1590
         MaxLength       =   30
         TabIndex        =   3
         Top             =   540
         Width           =   2535
      End
      Begin VB.Label lab��ˮ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "סԺ��(&N)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   540
         TabIndex        =   7
         Top             =   1005
         Width           =   945
      End
      Begin VB.Label lab������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�ϴ���(&M)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   540
         TabIndex        =   6
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label labҽ���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   540
         TabIndex        =   5
         Top             =   600
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3615
      TabIndex        =   1
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2400
      TabIndex        =   0
      Top             =   2550
      Width           =   1100
   End
End
Attribute VB_Name = "frmMain_�������󲡰���Ϣ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrWhere       As String

Public Property Get strWhere() As String
    strWhere = mstrWhere
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrH
    mstrWhere = ""
    If Not opt3.Value Then
        mstrWhere = mstrWhere & " And A.Up='" & IIf(opt1.Value, 1, 0) & "'"
    End If
    If txtҽ����.Text <> "" Then
        mstrWhere = mstrWhere & " And B.ҽ����='" & txtҽ����.Text & "'"
    End If
 
    If txt��ˮ��.Text <> "" Then
        mstrWhere = mstrWhere & " And A.RESIDENCE_NO='" & txt��ˮ��.Text & "'"
    End If
    If mstrWhere = "" Then
        MsgBox "����ѡ��һ��������", vbExclamation, gstrSysName
        Exit Sub
    End If
    Unload Me
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

