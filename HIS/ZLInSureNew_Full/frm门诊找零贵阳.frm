VERSION 5.00
Begin VB.Form frm����������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������Ϣ"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   Icon            =   "frm�����������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2520
      Top             =   1770
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ǰ��#�������շѵ���:"
      Height          =   1635
      Left            =   30
      TabIndex        =   2
      Top             =   90
      Width           =   4935
      Begin VB.TextBox txt������ 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1230
         Width           =   1455
      End
      Begin VB.TextBox txt�ֽ�֧�� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1230
         Width           =   1455
      End
      Begin VB.TextBox txt�󲡻��� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txt�����ܶ� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txt�����ʻ� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox txtҽ������ 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ֽ�֧��"
         Height          =   180
         Index           =   5
         Left            =   2460
         TabIndex        =   15
         Top             =   1297
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ܶ�"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ʋ���"
         Height          =   180
         Index           =   6
         Left            =   90
         TabIndex        =   10
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�󲡻���"
         Height          =   180
         Index           =   7
         Left            =   2460
         TabIndex        =   9
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Index           =   8
         Left            =   2460
         TabIndex        =   8
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʻ�"
         Height          =   180
         Index           =   9
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -90
      TabIndex        =   1
      Top             =   1770
      Width           =   5085
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   3540
      TabIndex        =   0
      Top             =   1890
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ע�⣺��������30���Ӻ��Զ��ر�"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   90
      TabIndex        =   12
      Top             =   1935
      Width           =   2700
   End
End
Attribute VB_Name = "frm�����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mdat As Date
Private Sub cmdOK_Click()
    If gblnLED And Val(txt�ֽ�֧��.Text) > 0 Then
        zl9LedVoice.Reset Nothing
        zl9LedVoice.Speak "#21" & Val(txt�ֽ�֧��.Text)
    End If
    Unload Me
End Sub
Public Sub ShowForm(ByVal intCOUNT As Long)
    Frame2.Caption = Replace(Frame2.Caption, "#", intCOUNT)
    txt�����ܶ�.Text = Format(g��������.dbl�����ܶ�, "0.00")
    txt�����ʻ�.Text = Format(g��������.dbl�����ʻ�, "0.00")
    txtҽ������.Text = Format(g��������.dblҽ������, "0.00")
    txtҽ�Ʋ���.Text = Format(g��������.dbl����Ա����, "0.00")
    txt�󲡻���.Text = Format(g��������.dbl�󲡻���, "0.00")
    txt�ֽ�֧��.Text = Format(g��������.dbl�ֽ�, "0.00")
    txt������.Text = Format(g��������.dbl������, "0.00")
    Label1.Caption = "ע�⣺��������" & mlngCloseTime & "���Ӻ��Զ��ر�"
    mdat = Now
    If gblnLED Then
        Call zl9LedVoice.DisplayBank("��������:" & intCOUNT & " �����ܶ�:" & txt�����ܶ�.Text, "�����ʻ�:" & txt�����ʻ�.Text & " ҽ������:" & txtҽ������.Text, _
                "ҽ�Ʋ���:" & txtҽ�Ʋ���.Text & " �󲡻���:" & txt�󲡻���.Text, "�ֽ�֧��:" & txt�ֽ�֧��.Text, "������:" & txt������.Text)
    End If
    Me.Show 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Timer1_Timer()
    '����ͣ��ʱ�䳬��30����ʱ���Զ��رձ����ڣ�Ҫ��ȻHIS����һֱ�޷��ύ������ȫԺ����
    If Abs(DateDiff("s", mdat, Now)) > mlngCloseTime Then Unload Me
End Sub

