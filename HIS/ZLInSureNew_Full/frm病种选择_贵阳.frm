VERSION 5.00
Begin VB.Form frm����ѡ��_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ICD��������ѡ��"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frm����ѡ��_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt�Ա� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   570
      Width           =   525
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1800
      TabIndex        =   6
      Top             =   1410
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3060
      TabIndex        =   7
      Top             =   1410
      Width           =   1100
   End
   Begin VB.TextBox txt������Ϣ 
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lbl�Ա� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      Height          =   180
      Left            =   660
      TabIndex        =   2
      Top             =   630
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   660
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ICD����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   390
      TabIndex        =   4
      Top             =   1020
      Width           =   630
   End
End
Attribute VB_Name = "frm����ѡ��_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrICD���� As String
Private mlng����ID As Long

Public Function ChooseDisease(ByVal lng����ID As Long) As String
    mlng����ID = lng����ID
    Me.Show 1
    ChooseDisease = mstrICD����
End Function

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    mstrICD���� = txt������Ϣ.Tag
    
    If Trim(mstrICD����) = "" Then
        MsgBox "����Ҫѡ��һ�ּ�����", vbInformation, gstrSysName
        txt������Ϣ.SetFocus
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select ����,�Ա�" & _
        " From ������Ϣ " & _
        " Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", mlng����ID)
    
    With rsTemp
        Me.txt���� = !����
        Me.txt�Ա� = !�Ա�
    End With
End Sub

Private Sub txt������Ϣ_GotFocus()
    Call zlControl.TxtSelAll(txt������Ϣ)
End Sub

Private Sub txt������Ϣ_KeyPress(KeyAscii As Integer)
    Dim strLike As String
    Dim StrInput As String
    Dim str�Ա� As String
    Dim blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If txt������Ϣ.Text = lbl������Ϣ.Tag And txt������Ϣ.Text <> "" Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf txt������Ϣ.Text = "" Then
        txt������Ϣ.Tag = "": lbl������Ϣ.Tag = ""
        Call zlCommFun.PressKey(vbKeyTab) '��������
    Else
        strLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
        StrInput = UCase(txt������Ϣ.Text)
        str�Ա� = txt�Ա�.Text
        If str�Ա� = "��" Then
            str�Ա� = " And (A.�Ա�����='��' Or A.�Ա����� is NULL)"
        ElseIf str�Ա� = "Ů" Then
            str�Ա� = " And (A.�Ա�����='Ů' Or A.�Ա����� is NULL)"
        Else
            str�Ա� = ""
        End If
        gstrSQL = "Select A.ID,A.����,A.����,A.����,A.����,A.˵��,A.�Ա�����,B.���" & _
            " From ��������Ŀ¼ A,����������� B" & _
            " Where A.���=B.���� And A.��� Not IN('B','Z')" & _
            " And (A.���� Like '" & StrInput & "%'" & _
            " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
            " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
            " Or Upper(A.����) Like '" & strLike & StrInput & "%')" & _
            " And Rownum<=100" & str�Ա� & _
            " Order by A.���,A.����"
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "��������Input", , , , , , True, _
            txt������Ϣ.Left + Me.Left, _
            txt������Ϣ.Top + Me.Top, txt������Ϣ.Height, blnCancel, , True)
        If Not rsTemp Is Nothing Then
            txt������Ϣ.Tag = rsTemp!����
            txt������Ϣ.Text = "(" & rsTemp!���� & ")" & rsTemp!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            txt������Ϣ.Tag = StrInput
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
