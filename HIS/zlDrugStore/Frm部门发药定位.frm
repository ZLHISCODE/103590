VERSION 5.00
Begin VB.Form Frm���ŷ�ҩ��λ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��λ"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "Frm���ŷ�ҩ��λ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtסԺ�� 
      Height          =   300
      Left            =   1140
      TabIndex        =   16
      Top             =   2940
      Width           =   2235
   End
   Begin VB.TextBox Txt����ҽ�� 
      Height          =   300
      Left            =   1140
      MaxLength       =   8
      TabIndex        =   5
      Top             =   990
      Width           =   1215
   End
   Begin VB.ComboBox Cob���� 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   2205
   End
   Begin VB.TextBox TxtNO 
      Height          =   300
      Left            =   1140
      MaxLength       =   8
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox Cob���� 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1380
      Width           =   2205
   End
   Begin VB.TextBox TxtҩƷ 
      Height          =   300
      Left            =   1140
      TabIndex        =   9
      Top             =   1770
      Width           =   1935
   End
   Begin VB.TextBox Txt���� 
      Height          =   300
      Left            =   1140
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Txt���� 
      Height          =   300
      Left            =   1140
      TabIndex        =   14
      Top             =   2550
      Width           =   1215
   End
   Begin VB.CommandButton cmdҩƷ 
      Caption         =   "��"
      Height          =   300
      Left            =   3060
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1770
      Width           =   285
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3780
      TabIndex        =   17
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3780
      TabIndex        =   18
      Top             =   690
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   3855
      Left            =   3540
      TabIndex        =   19
      Top             =   -120
      Width           =   45
   End
   Begin VB.Label lblסԺ�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "סԺ��(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   15
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label Lbl����ҽ�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ��(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   4
      Top             =   1050
      Width           =   630
   End
   Begin VB.Label Lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   270
      Width           =   630
   End
   Begin VB.Label LblNO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&NO"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   900
      TabIndex        =   2
      Top             =   660
      Width           =   180
   End
   Begin VB.Label Lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   6
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label LblҩƷ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   8
      Top             =   1830
      Width           =   630
   End
   Begin VB.Label Lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   11
      Top             =   2220
      Width           =   630
   End
   Begin VB.Label Lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   13
      Top             =   2610
      Width           =   630
   End
End
Attribute VB_Name = "Frm���ŷ�ҩ��λ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strReturn As String
Private lngҩ��ID As Long
Private mstrPrivs As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdOK_Click()
    '��֯���Ҵ�
    strReturn = ""
    If Cob����.ListIndex <> 0 Then
        strReturn = strReturn & IIf(strReturn = "", "", " And ") & _
        "����='" & Mid(Cob����.Text, InStr(1, Cob����.Text, "-") + 1) & "'"
    End If
    If Cob����.ListIndex <> 0 Then
        strReturn = strReturn & IIf(strReturn = "", "", " And ") & _
        "����='" & Mid(Cob����.Text, InStr(1, Cob����.Text, "-") + 1) & "'"
    End If
    If Trim(TxtNo) <> "" Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "NO='" & TxtNo.Text & "'"
    If Trim(Txt����ҽ��) <> "" Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "����ҽ�� Like '" & Txt����ҽ��.Text & "%'"
    If Val(txtҩƷ.Tag) <> 0 Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "ҩƷID=" & txtҩƷ.Tag
    If Trim(Txt����) <> "" Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "����='" & Txt����.Text & "'"
    If Trim(Txt����) <> "" Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "����='" & Txt����.Text & "'"
    If Trim(txtסԺ��) <> "" Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "סԺ��=" & txtסԺ��.Text
    
    If strReturn = "" Then
        MsgBox "��������Ҫ���ҵ����ݣ�", vbInformation, gstrSysName
        Cob����.SetFocus
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmdҩƷ_Click()
    Dim RecReturn As New ADODB.Recordset
    
'    With FrmҩƷѡ����
'        Set RecReturn = .ShowME(Me, 1, lngҩ��ID, , , False)
'    End With
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "ҩƷ���ŷ�ҩ", lngҩ��ID, lngҩ��ID)
    End If
    Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , lngҩ��ID, , , False, , , , , False)
        
    With RecReturn
        If .EOF Then Exit Sub
        txtҩƷ.Tag = !ҩƷID
        txtҩƷ = "[" & !ҩƷ���� & "]" & IIf(IsNull(!ͨ����), "", !ͨ����)
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsPart As New ADODB.Recordset
    strReturn = ""
    
    On Error GoTo errHandle
    Me.Txt����ҽ��.Enabled = IsHavePrivs(mstrPrivs, "ҽ����ѯ")
    
    With Cob����
        .Clear
        .AddItem "���е���"
        .AddItem "������ʵ�"
        .AddItem "סԺ���ʵ�"
        .AddItem "���ʱ�"
        .AddItem "ҽ��-����"
        .AddItem "ҽ��-����"
        .ListIndex = 0
    End With
    
    '���������÷�(�ٴ�������)
    Cob����.Clear
    Cob����.AddItem "���п���"
    
    gstrSQL = " Select ����||'-'||���� ����,ID From ���ű� " & _
             " Where ID in (Select ����ID From ��������˵�� Where �������� In ('�ٴ�','����') And ������� IN(2,3))" & _
             " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
             " Order By ����||'-'||���� "

    Set rsPart = zldatabase.OpenSQLRecord(gstrSQL, "Form_Load")
    With rsPart
        If .EOF Then
            MsgBox "���ʼ���ٴ����Ҽ��������ң������Ź���", vbInformation, gstrSysName
            Exit Sub
        End If
        Do While Not .EOF
            Cob����.AddItem !����
            .MoveNext
        Loop
        Cob����.ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ShowME(ByVal ҩ��ID_IN As Long, ByVal frmParent As Object, ByVal In_Ȩ�� As String) As String
    lngҩ��ID = ҩ��ID_IN
    mstrPrivs = In_Ȩ��
    
    Me.Show 1, frmParent
    ShowME = strReturn
End Function

Private Sub Form_Unload(Cancel As Integer)
        Call ReleaseSelectorRS
End Sub

Private Sub txtNO_GotFocus()
    Call SelAll(TxtNo)
End Sub

Private Sub TxtNO_Validate(Cancel As Boolean)
    Dim intYear As Integer, strYear As String
    If Trim(TxtNo) = "" Then Exit Sub
    '--���������λ,�򰴹������--
    Me.TxtNo = UCase(LTrim(Me.TxtNo))
    If Len(TxtNo) < 8 Then
        intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        Me.TxtNo = strYear & String(7 - Len(TxtNo), "0") & Me.TxtNo
    End If
End Sub

Private Sub txt����_GotFocus()
    Call SelAll(Txt����)
End Sub

Private Sub Txt����_GotFocus()
    Call SelAll(Txt����)
End Sub

Private Sub TxtҩƷ_GotFocus()
    Call SelAll(txtҩƷ)
End Sub

Private Sub TxtҩƷ_Validate(Cancel As Boolean)
    txtҩƷ = Trim(txtҩƷ)
    If txtҩƷ = "" Then
        txtҩƷ.Tag = 0
        Exit Sub
    End If
    
    Dim RecReturn As New ADODB.Recordset
    Dim sngLeft As Single, sngTop As Single
    
    If InStr(1, txtҩƷ, "[") <> 0 And InStr(1, txtҩƷ, "]") <> 0 Then txtҩƷ.Text = Mid(txtҩƷ.Text, 2, InStr(1, txtҩƷ, "]") - 2)
    sngLeft = Me.Left + txtҩƷ.Left + 50
    sngTop = Me.Top + (Me.Height - Me.ScaleHeight) + txtҩƷ.Top + txtҩƷ.Height - 100
    If DblFrmHeight + sngTop > Screen.Height Then sngTop = sngTop - DblFrmHeight - txtҩƷ.Height + 50
    
'    With FrmҩƷ��ѡѡ����
'        Set RecReturn = .ShowME(Me, 1, lngҩ��ID, , , TxtҩƷ.Text, sngLeft, sngTop, False)
'        If RecReturn.EOF Then Cancel = True: Exit Sub
'    End With
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "ҩƷ���ŷ�ҩ", lngҩ��ID, lngҩ��ID)
    End If
    Set RecReturn = frmSelector.ShowME(Me, 1, 2, UCase(txtҩƷ.Text), sngLeft, sngTop, lngҩ��ID, , , , False, , , , False)
    
    If RecReturn.EOF Then Cancel = True: Exit Sub
    txtҩƷ.Tag = RecReturn!ҩƷID
    txtҩƷ = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!ͨ����), "", RecReturn!ͨ����)
End Sub
