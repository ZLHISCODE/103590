VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDoubleBalanceFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtסԺ�� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3705
      MaxLength       =   18
      TabIndex        =   5
      Top             =   1290
      Width           =   1830
   End
   Begin VB.CheckBox chkDelRecord 
      Caption         =   "���˷ѽ����¼"
      Height          =   345
      Left            =   3705
      TabIndex        =   26
      Top             =   255
      Width           =   2160
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      IMEMode         =   1  'ON
      Left            =   1065
      MaxLength       =   64
      TabIndex        =   2
      Top             =   915
      Width           =   1830
   End
   Begin VB.TextBox txt����� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1065
      MaxLength       =   18
      TabIndex        =   4
      Top             =   1290
      Width           =   1830
   End
   Begin VB.OptionButton opt���� 
      Caption         =   "���ﲡ�˺�סԺ����"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   3705
      TabIndex        =   13
      Top             =   2535
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton opt���� 
      Caption         =   "סԺ����"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   2415
      TabIndex        =   12
      Top             =   2535
      Width           =   1020
   End
   Begin VB.OptionButton opt���� 
      Caption         =   "���ﲡ��"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   1065
      TabIndex        =   11
      Top             =   2535
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5880
      TabIndex        =   14
      Top             =   225
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5880
      TabIndex        =   15
      Top             =   645
      Width           =   1100
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   5880
      TabIndex        =   16
      Top             =   1605
      Width           =   1100
   End
   Begin VB.ComboBox cbo����Ա 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3705
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   915
      Width           =   1830
   End
   Begin VB.TextBox txtNOBegin 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1065
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1680
      Width           =   1830
   End
   Begin VB.TextBox txtNoEnd 
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3705
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1680
      Width           =   1830
   End
   Begin VB.TextBox txtFactBegin 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1065
      TabIndex        =   8
      Top             =   2055
      Width           =   1830
   End
   Begin VB.TextBox txtFactEnd 
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3705
      TabIndex        =   10
      Top             =   2055
      Width           =   1830
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   1065
      TabIndex        =   1
      Top             =   525
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   185925635
      CurrentDate     =   36588
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1065
      TabIndex        =   0
      Top             =   105
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   185925635
      CurrentDate     =   36588
   End
   Begin VB.Label lblסԺ�� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "סԺ��"
      Height          =   180
      Left            =   3090
      TabIndex        =   27
      Top             =   1350
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   630
      TabIndex        =   25
      Top             =   975
      Width           =   360
   End
   Begin VB.Label lbl����� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      Height          =   180
      Left            =   225
      TabIndex        =   24
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label lblFil 
      Alignment       =   1  'Right Justify
      Caption         =   "������Դ"
      Height          =   180
      Left            =   60
      TabIndex        =   23
      Top             =   2535
      Width           =   930
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼʱ��"
      Height          =   180
      Left            =   270
      TabIndex        =   22
      Top             =   165
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��"
      Height          =   180
      Left            =   270
      TabIndex        =   21
      Top             =   585
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   3255
      TabIndex        =   20
      Top             =   1740
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݺ�"
      Height          =   180
      Left            =   450
      TabIndex        =   19
      Top             =   1740
      Width           =   540
   End
   Begin VB.Label lbl����Ա 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����Ա"
      Height          =   180
      Left            =   3090
      TabIndex        =   18
      Top             =   975
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   3255
      TabIndex        =   17
      Top             =   2115
      Width           =   180
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ʊ�ݺ�"
      Height          =   180
      Left            =   450
      TabIndex        =   9
      Top             =   2115
      Width           =   540
   End
End
Attribute VB_Name = "frmDoubleBalanceFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModul As Long, mstrPrivs As String, mfrmParent As frmDoubleBalanceNormal
Public mblnInit As Boolean

Public Sub InitFilter(frmMain As Object, lngModul As Long, strPrivs As String)
    Set mfrmParent = frmMain
    mlngModul = lngModul
    mstrPrivs = strPrivs
    mblnInit = True
    Me.Show vbModal, frmMain
End Sub

Public Function FilterInited() As Boolean
    FilterInited = mblnInit
End Function

Private Sub cbo����Ա_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo����Ա.hWnd, KeyAscii)
        If lngIdx = -1 And cbo����Ա.ListCount > 0 Then lngIdx = 0
        cbo����Ա.ListIndex = lngIdx
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub

Private Sub cmdOK_Click()
    mblnInit = True
    Call mfrmParent.ReadData(0, mstrPrivs)
    Me.Hide
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub LoadOperator()
    Dim rsTmp As New ADODB.Recordset, i As Integer
    '����Ա
    cbo����Ա.Clear
    If InStr(mstrPrivs, "���в���Ա") > 0 Then
        cbo����Ա.AddItem "�����շ�Ա"
        Set rsTmp = GetPersonnel("�����շ�Ա", True)
        For i = 1 To rsTmp.RecordCount
            cbo����Ա.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo����Ա.ItemData(cbo����Ա.NewIndex) = rsTmp!ID
            If rsTmp!ID = UserInfo.ID Then cbo����Ա.ListIndex = cbo����Ա.NewIndex
            rsTmp.MoveNext
        Next
    Else
        cbo����Ա.AddItem UserInfo.���� & "-" & UserInfo.����
        cbo����Ա.ItemData(cbo����Ա.NewIndex) = UserInfo.ID
    End If
    If cbo����Ա.ListIndex = -1 And cbo����Ա.ListCount > 0 Then cbo����Ա.ListIndex = 0
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_Load()
    Dim Curdate As Date
    Call LoadOperator
    Curdate = zlDatabase.Currentdate
    dtpBegin.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = dtpBegin.MaxDate
    txt����.Text = "": txtFactBegin.Text = "": txtFactEnd.Text = ""
    txtNOBegin.Text = "": txtNoEnd.Text = ""
    txt�����.Text = "": txtסԺ��.Text = ""
    chkDelRecord.Value = 0
    Call opt����_Click(0)
End Sub

Private Sub opt����_Click(Index As Integer)
    If opt����(0).Value Then
        lbl�����.Visible = True
        txt�����.Visible = True
        lblסԺ��.Visible = False
        txtסԺ��.Visible = False
    ElseIf opt����(1).Value Then
        lbl�����.Visible = False
        txt�����.Visible = False
        lblסԺ��.Visible = True
        txtסԺ��.Visible = True
        lblסԺ��.Left = lbl�����.Left
        txtסԺ��.Left = txt�����.Left
    ElseIf opt����(2).Value Then
        lbl�����.Visible = True
        txt�����.Visible = True
        lblסԺ��.Visible = True
        txtסԺ��.Visible = True
        lblסԺ��.Left = lbl����Ա.Left
        txtסԺ��.Left = cbo����Ա.Left
    End If
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
    
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlControl.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 13)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 13)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtFactBegin_GotFocus()
    zlControl.TxtSelAll txtFactBegin
End Sub

Private Sub txtFactBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactEnd_GotFocus()
    zlControl.TxtSelAll txtFactEnd
End Sub

Private Sub txtFactEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactBegin_Change()
    txtFactEnd.Enabled = Not (Trim(txtFactBegin.Text) = "")
    If Trim(txtFactBegin.Text = "") Then txtFactEnd.Text = ""
End Sub
