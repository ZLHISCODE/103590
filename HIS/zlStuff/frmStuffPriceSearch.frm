VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStuffPriceSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra��Χ 
      Height          =   3810
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5385
      Begin VB.CheckBox ChkҩƷ 
         Caption         =   "����"
         Height          =   300
         Left            =   480
         TabIndex        =   22
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox TxtҩƷ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2400
         Width           =   3615
      End
      Begin VB.CommandButton CmdҩƷ 
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4920
         TabIndex        =   20
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox chkPrice 
         Caption         =   "���ۼ۵���"
         Height          =   300
         Left            =   480
         TabIndex        =   9
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CheckBox chkִ������ 
         Caption         =   "ִ������"
         Height          =   300
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox chk�������� 
         Caption         =   "��������"
         Height          =   300
         Left            =   480
         TabIndex        =   7
         Top             =   800
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txt����NO 
         Height          =   300
         Left            =   3690
         MaxLength       =   10
         TabIndex        =   6
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox txt��ʼNo 
         Height          =   300
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   5
         Top             =   360
         Width           =   1605
      End
      Begin VB.CheckBox chkPriceAndCost 
         Caption         =   "�ɱ��ۺ��ۼ�һ�����"
         Height          =   300
         Left            =   480
         TabIndex        =   4
         Top             =   3240
         Width           =   2100
      End
      Begin VB.CheckBox chkCost 
         Caption         =   "���ɱ��۵���"
         Height          =   300
         Left            =   2400
         TabIndex        =   3
         Top             =   2880
         Width           =   1400
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   196411395
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Index           =   0
         Left            =   3585
         TabIndex        =   11
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   196411395
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   12
         Top             =   1845
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   196411395
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Index           =   1
         Left            =   3585
         TabIndex        =   13
         Top             =   1845
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   196411395
         CurrentDate     =   36263
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   3345
         TabIndex        =   19
         Top             =   1140
         Width           =   180
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   0
         Left            =   900
         TabIndex        =   18
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   3345
         TabIndex        =   17
         Top             =   1905
         Width           =   180
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ִ������"
         Height          =   180
         Index           =   1
         Left            =   900
         TabIndex        =   16
         Top             =   1905
         Width           =   720
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   15
         Top             =   420
         Width           =   180
      End
      Begin VB.Label LblNO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���ۻ��ܺ�"
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   420
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5520
      TabIndex        =   1
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1100
   End
End
Attribute VB_Name = "frmStuffPriceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrResult As String

Private Type Type_Condition '����ʱ���õ�����
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    dateִ��ʱ�俪ʼ As Date
    dateִ��ʱ����� As Date
End Type

Private mSQLCondition As Type_Condition

Private Sub chkִ������_Click()
    If chkִ������.Value = 1 Then
        dtp��ʼʱ��(1).Enabled = True
        dtp����ʱ��(1).Enabled = True
    Else
        dtp��ʼʱ��(1).Enabled = False
        dtp����ʱ��(1).Enabled = False
    End If
End Sub

Private Sub ChkҩƷ_Click()
    If ChkҩƷ.Value = 1 Then
        TxtҩƷ.Enabled = True
        CmdҩƷ.Enabled = True
    Else
        TxtҩƷ.Enabled = False
        CmdҩƷ.Enabled = False
    End If
End Sub

Private Sub Cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim date����ʱ�俪ʼ As Date
    Dim date����ʱ����� As Date
    Dim dateִ��ʱ�俪ʼ As Date
    Dim dateִ��ʱ����� As Date
    
    mstrResult = ""
    If Trim(txt��ʼNo.Text) <> "" Then
        If IsNumeric(txt��ʼNo.Text) Then
            If Len(txt��ʼNo.Text) < 10 Then
                MsgBox "��������ȷ�ĵ��ۻ��ܺţ�ȫ����10λ����", vbInformation, gstrSysName
                Me.txt��ʼNo.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "��������ȷ�ĵ��ۻ��ܺţ�ȫ����10λ����", vbInformation, gstrSysName
            Me.txt��ʼNo.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txt����NO.Text) <> "" Then
        If CLng(Val(Trim(txt��ʼNo.Text))) > CLng(Val(Trim(txt����NO.Text))) Then
            MsgBox "��ʼ���ۻ��ܺŲ���С�ڽ������ۻ��ܺţ�", vbInformation, gstrSysName
            Me.txt��ʼNo.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txt����NO.Text) <> "" Then
        If IsNumeric(txt����NO.Text) Then
            If Len(txt����NO.Text) < 10 Then
                MsgBox "��������ȷ�ĵ��ۻ��ܺţ�ȫ����10λ����", vbInformation, gstrSysName
                Me.txt����NO.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "��������ȷ�ĵ��ۻ��ܺţ�ȫ����10λ����", vbInformation, gstrSysName
            Me.txt����NO.SetFocus
            Exit Sub
        End If
    End If
    
    If ChkҩƷ.Value = 1 Then
        If Val(TxtҩƷ.Tag) = 0 Then
            MsgBox "��ѡ�����ѯ��ҩƷ��Ϣ��", vbInformation, gstrSysName
            Me.TxtҩƷ.SetFocus
            Exit Sub
        End If
    End If
    
    '���ۻ��ܺ�
    If Trim(txt��ʼNo.Text) <> "" And Trim(txt����NO.Text <> "") Then
        mstrResult = " a.���ۺ� >= " & txt��ʼNo.Text & " and a.���ۺ� <= " & txt����NO.Text
    ElseIf Trim(txt��ʼNo.Text) <> "" And Trim(txt����NO.Text = "") Then
        mstrResult = " a.���ۺ� >= " & txt��ʼNo.Text
    ElseIf Trim(txt��ʼNo.Text) = "" And Trim(txt����NO.Text <> "") Then
        mstrResult = " a.���ۺ� <= " & txt����NO.Text
    End If
    
    '����
    If chk��������.Value = 1 And chkִ������.Value = 0 Then
        If mstrResult = "" Then
            mstrResult = " a.�������� between [1] and [2] "
        Else
            mstrResult = mstrResult + " and a.�������� between [1] and [2] "
        End If
    ElseIf chk��������.Value = 0 And chkִ������.Value = 1 Then
        If mstrResult = "" Then
            mstrResult = " a.ִ������ between [3] and [4] "
        Else
            mstrResult = mstrResult + " and a.ִ������ between [3] and [4] "
        End If
    ElseIf chk��������.Value = 1 And chkִ������.Value = 1 Then
        If mstrResult = "" Then
            mstrResult = " a.�������� between [1] and [2] and a.ִ������ between [3] and [4] "
        Else
            mstrResult = mstrResult + " and a.�������� between [1] and [2] and a.ִ������ between [3] and [4] "
        End If
    End If
    '��������
    If chk��������.Value = 1 Then
        mSQLCondition.date����ʱ�俪ʼ = CDate(Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & " 00:00:00")
        mSQLCondition.date����ʱ����� = CDate(Format(dtp����ʱ��(0), "yyyy-mm-dd") & " 23:59:59")
    End If
    
    If chkִ������.Value = 1 Then
        mSQLCondition.dateִ��ʱ�俪ʼ = CDate(Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & " 00:00:00")
        mSQLCondition.dateִ��ʱ����� = CDate(Format(dtp����ʱ��(1), "yyyy-mm-dd") & " 23:59:59")
    End If
    'ִ������
    
    '��������
    If chkPrice.Value = 1 And chkCost.Value = 0 And chkPriceAndCost.Value = 0 Then '�����ۼ�
        If mstrResult = "" Then
            mstrResult = " a.����=0 "
        Else
            mstrResult = mstrResult + " and a.����=0 "
        End If
    ElseIf chkPrice.Value = 0 And chkCost.Value = 1 And chkPriceAndCost.Value = 0 Then '�����ɱ���
        If mstrResult = "" Then
            mstrResult = " a.����=1 "
        Else
            mstrResult = mstrResult + " and a.����=1 "
        End If
    ElseIf chkPrice.Value = 0 And chkCost.Value = 0 And chkPriceAndCost.Value = 1 Then '�ɱ����ۼ�һ�����
        If mstrResult = "" Then
            mstrResult = " a.����=2 "
        Else
            mstrResult = mstrResult + " and a.����=2 "
        End If
    ElseIf chkPrice.Value = 1 And chkCost.Value = 1 And chkPriceAndCost.Value = 0 Then '�����ۼۺͽ����ɱ���
        If mstrResult = "" Then
            mstrResult = " (a.����=0 or a.����=1) "
        Else
            mstrResult = mstrResult + " and (a.����=0 or a.����=1) "
        End If
    ElseIf chkPrice.Value = 1 And chkCost.Value = 0 And chkPriceAndCost.Value = 1 Then '�����ۼۺͳɱ����ۼ�һ�����
        If mstrResult = "" Then
            mstrResult = " (a.����=0 or a.����=2) "
        Else
            mstrResult = mstrResult + " and (a.����=0 or a.����=2) "
        End If
    ElseIf chkPrice.Value = 0 And chkCost.Value = 1 And chkPriceAndCost.Value = 1 Then '�����ɱ��ۺͳɱ����ۼ�һ�����
        If mstrResult = "" Then
            mstrResult = " (a.����=1 or a.����=2) "
        Else
            mstrResult = mstrResult + " and (a.����=1 or a.����=2) "
        End If
    End If
    
    'ҩƷ
    If Val(TxtҩƷ.Tag) <> 0 Then
        If mstrResult = "" Then
            mstrResult = " a.���ۺ� In (Select ���ۻ��ܺ� From �շѼ�Ŀ Where �շ�ϸĿid = " & TxtҩƷ.Tag & GetPriceClassString("") & _
                            " union all " & _
                         " Select  ���ۻ��ܺ� From �ɱ��۵�����Ϣ Where ҩƷid = " & TxtҩƷ.Tag & ")"
        Else
            mstrResult = mstrResult & " and a.���ۺ� In (Select ���ۻ��ܺ� From �շѼ�Ŀ Where �շ�ϸĿid =" & TxtҩƷ.Tag & GetPriceClassString("") & _
                        " union all " & _
                         " Select  ���ۻ��ܺ� From �ɱ��۵�����Ϣ Where ҩƷid = " & TxtҩƷ.Tag & ")"
        End If
    End If
    Unload Me
End Sub

Private Sub CmdҩƷ_Click()
    Dim RecReturn As Recordset
    
    On Error GoTo ErrHandle
    
    Set RecReturn = Frm����ѡ����.ShowMe(Me, 1, 0)
    If RecReturn.RecordCount = 0 Then Exit Sub
    TxtҩƷ = "[" & RecReturn!���� & "]" & RecReturn!����
    TxtҩƷ.Tag = RecReturn!����ID
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Me.dtp����ʱ��(0) = Sys.Currentdate
    Me.dtp����ʱ��(1) = Me.dtp����ʱ��(0)
    Me.dtp��ʼʱ��(0) = DateAdd("d", -7, Me.dtp����ʱ��(0))
    Me.dtp��ʼʱ��(1) = Me.dtp��ʼʱ��(0)
End Sub

Public Sub ShowMe(ByVal FrmParent As Form, ByRef strResult As String, ByRef date����ʱ�俪ʼ As Date, ByRef date����ʱ����� As Date, ByRef date���ʱ�俪ʼ As Date, ByRef date���ʱ����� As Date)
    Me.Show vbModal, FrmParent
        
    strResult = mstrResult
    date����ʱ�俪ʼ = mSQLCondition.date����ʱ�俪ʼ
    date����ʱ����� = mSQLCondition.date����ʱ�����
    date���ʱ�俪ʼ = mSQLCondition.dateִ��ʱ�俪ʼ
    date���ʱ����� = mSQLCondition.dateִ��ʱ�����
End Sub

Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt��ʼNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(TxtҩƷ.Text) = "" Then Exit Sub
    sngLeft = TxtҩƷ.Left
    sngTop = Me.Top + TxtҩƷ.Height + TxtҩƷ.Top '  50
    If sngTop + 4530 > Screen.Height Then
        sngTop = sngTop - TxtҩƷ.Height - 4530
    End If
    
    strKey = Trim(TxtҩƷ.Text)
    If Mid(strKey, 1, 1) = "[" Then
        If InStr(2, strKey, "]") <> 0 Then
            strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
        Else
            strKey = Mid(strKey, 2)
        End If
    End If
    
    Set RecReturn = FrmMulitSel.ShowSelect(Me, 1, , , , strKey, sngLeft, sngTop)
    If RecReturn.RecordCount = 0 Then Exit Sub
    TxtҩƷ = "[" & RecReturn!���� & "]" & RecReturn!����
    TxtҩƷ.Tag = RecReturn!����ID
    
    TxtҩƷ.SetFocus
End Sub

