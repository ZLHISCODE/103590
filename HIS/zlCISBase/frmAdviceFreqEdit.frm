VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceFreqEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ��Ƶ�ʱ༭"
   ClientHeight    =   3015
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4830
   Icon            =   "frmAdviceFreqEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComCtl2.UpDown UDƵ�ʼ�� 
      Height          =   300
      Left            =   2835
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2145
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtEdit(5)"
      BuddyDispid     =   196611
      BuddyIndex      =   5
      OrigLeft        =   2836
      OrigTop         =   2145
      OrigRight       =   3076
      OrigBottom      =   2445
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UDƵ�ʴ��� 
      Height          =   300
      Left            =   2835
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1740
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtEdit(4)"
      BuddyDispid     =   196611
      BuddyIndex      =   4
      OrigLeft        =   2836
      OrigTop         =   1743
      OrigRight       =   3076
      OrigBottom      =   2043
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1341
      Width           =   1785
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   1
      Top             =   939
      Width           =   1785
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   3570
      TabIndex        =   8
      Top             =   2295
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3510
      Left            =   3345
      TabIndex        =   15
      Top             =   -300
      Width           =   30
   End
   Begin VB.ComboBox cbo�����λ 
      Height          =   300
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmAdviceFreqEdit.frx":000C
      Left            =   1290
      List            =   "frmAdviceFreqEdit.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2550
      Width           =   1785
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   1290
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "1"
      Top             =   2145
      Width           =   1530
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1290
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "1"
      Top             =   1743
      Width           =   1530
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1290
      MaxLength       =   20
      TabIndex        =   0
      Top             =   537
      Width           =   1785
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1290
      MaxLength       =   3
      TabIndex        =   10
      Top             =   135
      Width           =   1785
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   7
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3570
      TabIndex        =   6
      Top             =   390
      Width           =   1100
   End
   Begin VB.Label lblӢ������ 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӣ������(&E)"
      Height          =   180
      Left            =   225
      TabIndex        =   17
      Top             =   1395
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&S)"
      Height          =   180
      Left            =   585
      TabIndex        =   16
      Top             =   1005
      Width           =   630
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����λ(&U)"
      Height          =   180
      Left            =   225
      TabIndex        =   14
      Top             =   2610
      Width           =   990
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ƶ�ʼ��(&J)"
      Height          =   180
      Left            =   225
      TabIndex        =   13
      Top             =   2205
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ƶ�ʴ���(&M)"
      Height          =   180
      Left            =   225
      TabIndex        =   12
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   585
      TabIndex        =   11
      Top             =   600
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&B)"
      Height          =   180
      Left            =   585
      TabIndex        =   9
      Top             =   195
      Width           =   630
   End
End
Attribute VB_Name = "frmAdviceFreqEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrCode As String
Public mbytType As Byte '1-��ҽ,2-��ҽ

Private mstr�����λ As String
Private mintƵ�ʴ��� As Integer
Private mintƵ�ʼ�� As Integer

Private mblnChange As Boolean

Private Sub cbo�����λ_Click()
    If Visible Then mblnChange = True
    
    If cbo�����λ.Text = "��" Then
        txtEdit(5).Enabled = False
        UDƵ�ʼ��.Enabled = False
        txtEdit(5).Text = 1
    ElseIf cbo�����λ.Text = "����" Then
        txtEdit(4).Enabled = False
        UDƵ�ʴ���.Enabled = False
        txtEdit(4).Text = 1
    Else
        txtEdit(5).Enabled = True
        UDƵ�ʼ��.Enabled = True
        txtEdit(4).Enabled = True
        UDƵ�ʴ���.Enabled = True
    End If
    
    If cbo�����λ.Text = "����" Then
        txtEdit(5).MaxLength = 3
        UDƵ�ʼ��.Max = 999
    Else
        txtEdit(5).MaxLength = 2
        UDƵ�ʼ��.Max = 99
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim strDel As String
    Dim blnDel As Boolean
    
    If txtEdit(0).Text = "" Then
        MsgBox "����������롣", vbInformation, gstrSysName
        txtEdit(0).SetFocus: Exit Sub
    End If
    
    If txtEdit(1).Text = "" Then
        MsgBox "�����������ơ�", vbInformation, gstrSysName
        txtEdit(1).SetFocus: Exit Sub
    End If
    
    strSql = "Select 1 From ����Ƶ����Ŀ Where  Nvl(���÷�Χ, 0) <> 1 And Nvl(���÷�Χ, 0) <> 2 And ����=[1] and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "cmdOK_Click", txtEdit(1).Text)
    
    If rsTmp.RecordCount > 0 Then
        MsgBox "�̶�Ƶ�����Ѿ�������ͬ���Ƶ�Ƶ�ʡ�", vbInformation, gstrSysName
        txtEdit(1).SetFocus: Exit Sub
    End If
    
    If zlCommFun.ActualLen(txtEdit(1).Text) > txtEdit(1).MaxLength Then
        MsgBox "�������ݹ��������ֻ�ܰ���" & txtEdit(1).MaxLength & "���ַ���" & txtEdit(1).MaxLength \ 2 & "�����֡�", vbInformation, gstrSysName
        txtEdit(1).SetFocus: Exit Sub
    End If
    
'    If Val(txtEdit(4).Text) <> 1 And Val(txtEdit(5).Text) <> 1 Then
'        MsgBox "Ƶ�ʴ�����Ƶ�ʼ������Ӧ����һ��Ϊ 1 ��", vbInformation, gstrSysName
'        txtEdit(5).SetFocus: Exit Sub
'    End If
        
    If mstrCode = "" Then
        strSql = "ZL_����Ƶ����Ŀ_Insert('" & txtEdit(0).Text & "','" & txtEdit(1).Text & "','" & txtEdit(2).Text & "','" & txtEdit(3).Text & "'," & txtEdit(4).Text & "," & txtEdit(5).Text & ",'" & cbo�����λ.Text & "'," & mbytType & ")"
    Else
        If Val(txtEdit(4).Text) <> mintƵ�ʴ��� Then
            If MsgBox("�������Ƶ�ʴ������⽫�����Ƶ����Ŀ���е�ʱ�����á�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            blnDel = True
        ElseIf Val(txtEdit(5).Text) <> mintƵ�ʼ�� And cbo�����λ.Text <> "����" Then
            If MsgBox("�������Ƶ�ʼ�����⽫�����Ƶ����Ŀ���е�ʱ�����á�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            blnDel = True
        ElseIf cbo�����λ.Text <> mstr�����λ Then
            If MsgBox("������˼����λ���⽫�����Ƶ����Ŀ���е�ʱ�����á�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            blnDel = True
        End If
        If blnDel Then strDel = "ZL_����Ƶ��ʱ��_Delete('" & mstrCode & "')"
        strSql = "ZL_����Ƶ����Ŀ_Update('" & mstrCode & "','" & txtEdit(0).Text & "','" & txtEdit(1).Text & "','" & txtEdit(2).Text & "','" & txtEdit(3).Text & "'," & txtEdit(4).Text & "," & txtEdit(5).Text & ",'" & cbo�����λ.Text & "')"
    End If
        
    On Error GoTo errH
    gcnOracle.BeginTrans
    If strDel <> "" Then
        Call zlDatabase.ExecuteProcedure(strDel, Me.Caption)
    End If
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    If mstrCode <> "" Then
        mblnChange = False
        gblnOK = True
        Unload Me
    Else
        Call frmAdviceFreq.LoadItems("_" & txtEdit(0).Text)
        Call Form_Load
        mblnChange = False
        gblnOK = True
        txtEdit(1).SetFocus
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    gblnOK = False
    mblnChange = False
    
    On Error GoTo errH
    
    If mstrCode <> "" Then
        '�޸�
        strSql = "Select * From ����Ƶ����Ŀ Where ����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrCode)
                
        cbo�����λ.Text = Nvl(rsTmp!�����λ)
        txtEdit(0).Text = Nvl(rsTmp!����)
        txtEdit(1).Text = Nvl(rsTmp!����)
        txtEdit(2).Text = Nvl(rsTmp!����)
        txtEdit(3).Text = Nvl(rsTmp!Ӣ������)
        txtEdit(4).Text = Nvl(rsTmp!Ƶ�ʴ���)
        txtEdit(5).Text = Nvl(rsTmp!Ƶ�ʼ��)
        
        mstr�����λ = Nvl(rsTmp!�����λ)
        mintƵ�ʴ��� = Nvl(rsTmp!Ƶ�ʴ���)
        mintƵ�ʼ�� = Nvl(rsTmp!Ƶ�ʼ��)
    Else
        '����
        txtEdit(0).Text = ""
        txtEdit(1).Text = ""
        txtEdit(2).Text = ""
        txtEdit(3).Text = ""
        
        strSql = "Select ZL_IncStr(Max(����)) as ���� From ����Ƶ����Ŀ"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!����) Then
                txtEdit(0).Text = rsTmp!����
            End If
        End If
        If txtEdit(0).Text = "" Then
            txtEdit(0).Text = String(txtEdit(0).MaxLength - 1, "0") & "1"
        End If
        If cbo�����λ.ListIndex = -1 Then cbo�����λ.Text = "��"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("�����޸���������ݣ�ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
    
    mstrCode = ""
    mbytType = 0
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Visible Then mblnChange = True
    '��������
    If Index = 1 And Visible Then txtEdit(2).Text = zlCommFun.SpellCode(txtEdit(Index).Text)
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If txtEdit(Index).IMEMode = 0 Then Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If txtEdit(Index).IMEMode = 0 Then Call zlCommFun.OpenIme
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If InStr("-", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    ElseIf Index = 4 Or Index = 5 Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If Not zlCommFun.StrIsValid(txtEdit(Index).Text, txtEdit(Index).MaxLength) Then
        Cancel = True
    ElseIf Index = 4 Or Index = 5 Then
        If Not IsNumeric(txtEdit(Index).Text) Or Val(txtEdit(Index).Text) <= 0 Then
            Cancel = True
            MsgBox "�����������Ҵ����㣡"
        End If
    End If
End Sub
