VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAutoArchive 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�Զ���������"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame frmAutoArchive 
      Caption         =   "ʱ���������"
      Height          =   2415
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   3975
      Begin VB.TextBox txtDay 
         Height          =   300
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "1"
         Top             =   360
         Width           =   1275
      End
      Begin VB.OptionButton optDay 
         Caption         =   "ÿ��"
         Height          =   300
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "ÿ��"
         Height          =   300
         Left            =   360
         TabIndex        =   8
         Top             =   890
         Width           =   1095
      End
      Begin VB.TextBox txtMonth 
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "1"
         Top             =   870
         Width           =   1455
      End
      Begin VB.ComboBox cobTimeArchiveStyle 
         Height          =   315
         ItemData        =   "frmAutoArchive.frx":0000
         Left            =   1440
         List            =   "frmAutoArchive.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Top             =   1350
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16646146
         CurrentDate     =   38226
      End
      Begin MSComCtl2.UpDown udMonth 
         Height          =   300
         Left            =   2880
         TabIndex        =   9
         Top             =   870
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udDay 
         Height          =   300
         Left            =   660
         TabIndex        =   15
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "���                 ��"
         Height          =   180
         Left            =   1440
         TabIndex        =   16
         Top             =   420
         Width           =   2070
      End
      Begin VB.Label Label2 
         Caption         =   "��                            ��"
         Height          =   195
         Left            =   3330
         TabIndex        =   12
         Top             =   930
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�鵵ʱ��"
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�鵵��ʽ"
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   1860
         Width           =   720
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   3600
      TabIndex        =   3
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ӧ��"
      Height          =   350
      Left            =   1920
      TabIndex        =   2
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   350
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CheckBox chkAutoArchive 
      Caption         =   "�����Զ��鵵(���Զ��鵵���м�¼)"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   4425
   End
End
Attribute VB_Name = "frmAutoArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkAutoArchive_Click()
    bAutoArchive = IIf(Me.chkAutoArchive.Value = 1, True, False)
    If bAutoArchive = False Then
        Me.frmAutoArchive.Enabled = False
    Else
        Me.frmAutoArchive.Enabled = True
    End If
End Sub


Private Sub Command1_Click()
    subApply
    Unload Me
End Sub

Private Sub Command2_Click()
    subApply
End Sub
Private Sub subApply()
    '�����в��Ա��浽��ʱ����
    If Me.chkAutoArchive = 1 Then       '�������Զ��鵵����
        bAutoArchive = True
        '����ʱ��鵵����
        If Me.optDay.Value = True Then
            strTimePolicy = "time,day," & Me.txtDay.Text & "," & Me.dtpTime.Hour & ":" & _
                  Me.dtpTime.Minute & ":" & Me.dtpTime.Second & "," & _
                  Me.cobTimeArchiveStyle.ListIndex & ",1"
        Else
            strTimePolicy = "time,month," & Me.txtMonth.Text & "," & Me.dtpTime.Hour & ":" & _
                  Me.dtpTime.Minute & ":" & Me.dtpTime.Second & "," & _
                  Me.cobTimeArchiveStyle.ListIndex & ",1"
        End If
    Else                        '����Ϊû���Զ��鵵����
        bAutoArchive = False
        strTimePolicy = "time,N/A"
    End If
    '����ʱ�������ݱ��浽ע���
    SaveSetting "ZLSOFT", "����ģ��\�鵵����", "ʱ��鵵����", strTimePolicy
    SaveSetting "ZLSOFT", "����ģ��\�鵵����", "ʹ���Զ��鵵", CStr(bAutoArchive)
End Sub
Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strTempPolicy() As String   '�ݴ汻�����������Զ��鵵����
    Dim strTempTime() As String     '�ݴ汻�����������Զ��鵵�����еĹ鵵ʱ��
    
    If bAutoArchive = True Then    '���Զ��鵵����
        Me.chkAutoArchive.Value = 1
        '����ʱ�����
        strTempPolicy = Split(strTimePolicy, ",")
        If UCase(strTempPolicy(1)) = "DAY" Then
            Me.optDay.Value = True
            Me.txtDay.Text = strTempPolicy(2)
        ElseIf UCase(strTempPolicy(1)) = "MONTH" Then
            Me.optMonth.Value = True
            Me.txtMonth.Text = strTempPolicy(2)
        End If
        strTempTime = Split(strTempPolicy(3), ":")
        Me.dtpTime.Hour = strTempTime(0)
        Me.dtpTime.Minute = strTempTime(1)
        Me.dtpTime.Second = strTempTime(2)
        If strTempPolicy(4) = "1" And strTempPolicy(5) = "1" Then
            Me.cobTimeArchiveStyle.ListIndex = 1        'ɾ���ҹ鵵
        ElseIf strTempPolicy(4) = "0" Then
            Me.cobTimeArchiveStyle.ListIndex = 0        'ֻ�鵵
        End If
    Else    'û���Զ��鵵����
        Me.chkAutoArchive.Value = 0
    End If
End Sub

Private Sub udDay_DownClick()
    Me.txtDay.Text = Val(Me.txtDay.Text) - 1
    If Val(Me.txtDay.Text) < 1 Then Me.txtDay.Text = 31
End Sub

Private Sub udDay_UpClick()
    Me.txtDay.Text = Val(Me.txtDay.Text) + 1
    If Val(Me.txtDay.Text) > 31 Then Me.txtDay.Text = 1
End Sub

Private Sub udMonth_DownClick()
    Me.txtMonth.Text = Val(Me.txtMonth.Text) - 1
    If Val(Me.txtMonth.Text) < 1 Then Me.txtMonth.Text = 31
End Sub

Private Sub udMonth_UpClick()
    Me.txtMonth.Text = Val(Me.txtMonth.Text) + 1
    If Val(Me.txtMonth.Text) > 31 Then Me.txtMonth.Text = 1
End Sub
