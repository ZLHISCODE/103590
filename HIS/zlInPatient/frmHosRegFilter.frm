VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHosRegFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   1935
      TabIndex        =   9
      Top             =   2325
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   195
      TabIndex        =   10
      Top             =   60
      Width           =   5760
      Begin VB.TextBox txt����� 
         Height          =   300
         Left            =   990
         TabIndex        =   6
         Top             =   1620
         Width           =   2085
      End
      Begin VB.TextBox txtסԺ��E 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3510
         TabIndex        =   5
         Top             =   1200
         Width           =   2085
      End
      Begin VB.TextBox txtסԺ��B 
         Height          =   300
         Left            =   990
         TabIndex        =   4
         Top             =   1200
         Width           =   2085
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   990
         TabIndex        =   2
         Top             =   780
         Width           =   2085
      End
      Begin VB.ComboBox cbo�Ǽ�Ա 
         Height          =   300
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   780
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker dtp��ԺE 
         Height          =   300
         Left            =   3495
         TabIndex        =   1
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   103743491
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp��ԺB 
         Height          =   300
         Left            =   990
         TabIndex        =   0
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   103743491
         CurrentDate     =   40544
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   390
         TabIndex        =   17
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   390
         TabIndex        =   16
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3210
         TabIndex        =   15
         Top             =   1260
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   210
         TabIndex        =   14
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3210
         TabIndex        =   13
         Top             =   420
         Width           =   180
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         Height          =   180
         Left            =   390
         TabIndex        =   12
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ�Ա"
         Height          =   180
         Left            =   3150
         TabIndex        =   11
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3105
      TabIndex        =   7
      Top             =   2325
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4290
      TabIndex        =   8
      Top             =   2325
      Width           =   1100
   End
End
Attribute VB_Name = "frmHosRegFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mbytType As Byte '��:�����嵥����
Public mstrFilter As String '��:����
Public mcllFilter As Collection

Private Sub cbo�Ǽ�Ա_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo�Ǽ�Ա.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo�Ǽ�Ա.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo�Ǽ�Ա.ListIndex = lngIdx
    If cbo�Ǽ�Ա.ListIndex = -1 And cbo�Ǽ�Ա.ListCount <> 0 Then cbo�Ǽ�Ա.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub

Private Sub cmdOK_Click()
    If IsNumeric(txtסԺ��E.Text) And IsNumeric(txtסԺ��B.Text) Then
        If CLng(txtסԺ��E.Text) <= CLng(txtסԺ��B.Text) Then
            MsgBox "��ʼסԺ��Ӧ��С�ڽ���סԺ�ţ�", vbInformation, gstrSysName
            txtסԺ��B.SetFocus: Exit Sub
        End If
    End If
    Call MakeFilter
    gblnOK = True
    Hide
End Sub

Private Sub dtp��ԺE_Change()
    dtp��ԺB.MaxDate = dtp��ԺE.Value
End Sub

Private Sub Form_Activate()
    dtp��ԺB.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim Curdate As Date, i As Integer
    
    txtסԺ��B.Text = ""
    txtסԺ��E.Text = ""
    txt�����.Text = ""
    
    '���ó�ʼ����(һ������Ժ)
    Curdate = zlDatabase.Currentdate
    dtp��ԺB.Value = Format(DateAdd("d", -7, Curdate), "yyyy-MM-dd 00:00:00")
    dtp��ԺE.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")

    cbo�Ǽ�Ա.Clear
    cbo�Ǽ�Ա.AddItem "���еǼ�Ա"
    cbo�Ǽ�Ա.ListIndex = 0
    
    Set rsTmp = GetPersonnel("��Ժ�Ǽ�Ա", True)
    For i = 1 To rsTmp.RecordCount
        cbo�Ǽ�Ա.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ID = UserInfo.ID Then cbo�Ǽ�Ա.ListIndex = cbo�Ǽ�Ա.NewIndex
        rsTmp.MoveNext
    Next
End Sub

Private Sub MakeFilter()
    
    'by lesfeng 2010-1-11 �����Ż�
    Set mcllFilter = New Collection
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "��Ժ����"
    mcllFilter.Add Array("", ""), "סԺ��"
    mcllFilter.Add "", "��������"
    mcllFilter.Add "", "�Ǽ���"
    mcllFilter.Add "", "�����"
    
'    mstrFilter = ""
'    mstrFilter = mstrFilter & " And B.��Ժ���� Between To_Date('" & Format(dtp��ԺB.Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp��ԺE.Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    
    mstrFilter = ""
    mstrFilter = mstrFilter & " And (B.��Ժ����  Between [1] And [2]) "
    mcllFilter.Remove "��Ժ����"
    mcllFilter.Add Array(Format(dtp��ԺB, "yyyy-MM-dd hh:mm:ss"), Format(dtp��ԺE, "yyyy-MM-dd hh:mm:ss")), "��Ժ����"
          
    If IsNumeric(txtסԺ��B.Text) And IsNumeric(txtסԺ��E.Text) Then
        mstrFilter = mstrFilter & " And A.����ID In (Select Distinct ����ID From ������ҳ Where סԺ�� Between [3] And [4]) "
    ElseIf IsNumeric(txtסԺ��B.Text) Then
        mstrFilter = mstrFilter & " And A.����ID = (Select Nvl(Max(����ID),0) as ����ID From ������ҳ Where סԺ��=[3]) "
    End If
    
    mcllFilter.Remove "סԺ��"
    mcllFilter.Add Array(Trim(txtסԺ��B.Text), Trim(txtסԺ��E.Text)), "סԺ��"
    

    If cbo�Ǽ�Ա.ListIndex <> 0 Then
        mstrFilter = mstrFilter & " And B.�Ǽ���=[5]"
    End If
    
    '����17122 by lesfeng 2010-02-02
    If Trim(txt����.Text) <> "" Then
        mstrFilter = mstrFilter & " And NVL(B.����,A.����) like [7]"
    End If
    
    mcllFilter.Remove "��������"
    mcllFilter.Add Trim(txt����.Text), "��������"
    
    mcllFilter.Remove "�Ǽ���"
    mcllFilter.Add zlCommFun.GetNeedName(cbo�Ǽ�Ա.Text), "�Ǽ���"
    
    If IsNumeric(txt�����.Text) Then mstrFilter = mstrFilter & " And A.����� = [8] "
    mcllFilter.Remove "�����"
    mcllFilter.Add Trim(txt�����.Text), "�����"
    
End Sub

Private Sub txt�����_GotFocus()
    zlControl.TxtSelAll txt�����
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

'����17122 by lesfeng 2010-02-02
Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub
'����17122 by lesfeng 2010-02-02
Private Sub txt����_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtסԺ��B_Change()
    txtסԺ��E.Enabled = (Trim(txtסԺ��B.Text) <> "")
    If Not txtסԺ��E.Enabled Then txtסԺ��E.Text = ""
End Sub

Private Sub txtסԺ��B_GotFocus()
    zlControl.TxtSelAll txtסԺ��B
End Sub

Private Sub txtסԺ��B_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtסԺ��E_GotFocus()
    zlControl.TxtSelAll txtסԺ��E
End Sub

Private Sub txtסԺ��E_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
