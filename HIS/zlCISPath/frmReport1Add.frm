VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmReport1Add 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ���ӱ���"
   ClientHeight    =   1710
   ClientLeft      =   2760
   ClientTop       =   3720
   ClientWidth     =   4470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1100
   End
   Begin VB.Frame fraDate 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2950
      Begin MSMask.MaskEdBox txtScope 
         Height          =   375
         Left            =   1275
         TabIndex        =   0
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm"
         Mask            =   "####-##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1275
         TabIndex        =   2
         Top             =   1140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   102694915
         CurrentDate     =   40240
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1275
         TabIndex        =   1
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   102694915
         CurrentDate     =   40240
      End
      Begin VB.Label lbl�ڼ� 
         Caption         =   "ͳ���ڼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
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
         Left            =   240
         TabIndex        =   6
         Top             =   780
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmReport1Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public mlng·��ID As Long   'in ����
Public mstr�ڼ� As String   'out����
Public mblnOK As Boolean    'out

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, str�ڼ� As String
    Dim rstmp As ADODB.Recordset
    
    On Error GoTo errH
    If Not IsDate(txtScope.Text) Then
        Call MsgBox("������ڼ䲻���Ϲ淶��Ҫ����λ���+��λ�����֡�", vbInformation, gstrSysName)
        txtScope.SetFocus
        Exit Sub
    End If
    
    str�ڼ� = Format(txtScope.Text, "yyyymm")
    strSQL = "Select 1 From ·�������ļ� Where �ڼ� = [1] And ·��ID = [2]"
    Set rstmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, str�ڼ�, mlng·��ID)
    If rstmp.RecordCount > 0 Then
        Call MsgBox("���ڼ��Ѵ��ڣ�����������һ���ڼ䡣", vbInformation, gstrSysName)
        Exit Sub
    End If
    
    strSQL = "Zl_·�������ļ�_Insert(1," & str�ڼ� & ",To_date('" & dtpBegin.Value & "','yyyy-mm-dd')" & _
            ",To_date('" & dtpEnd.Value & "','yyyy-mm-dd')," & mlng·��ID & ",'" & UserInfo.���� & "')"
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
            
    mstr�ڼ� = str�ڼ�
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlcommfun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim Datsys As Date
    
    Datsys = zldatabase.Currentdate
    txtScope.Text = Format(Datsys, "yyyy-mm")
    dtpBegin.Value = Format(Datsys, "yyyy-mm-01")
    dtpEnd.Value = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(Datsys, "yyyy-mm-01")))))
    
End Sub
