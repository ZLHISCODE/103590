VERSION 5.00
Begin VB.Form frm����ѡ��_��Ϫũҽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frm����ѡ��_��Ϫũҽ.frx":0000
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
      Caption         =   "������Ϣ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   300
      TabIndex        =   4
      Top             =   1020
      Width           =   720
   End
End
Attribute VB_Name = "frm����ѡ��_��Ϫũҽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mlng����ID As Long

Public Function ChooseDisease(ByVal lng����ID As Long) As Long
    mlng����ID = lng����ID
    Me.Show 1
    ChooseDisease = mlng����ID
End Function

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    mlng����ID = Val(txt������Ϣ.Tag)
    
    If mlng����ID = 0 Then
        MsgBox "����Ҫѡ��һ�ּ�����", vbInformation, gstrSysName
        txt������Ϣ.SetFocus
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select C.����,C.�Ա�,B.ID AS ����ID,'('||B.����||')'||B.���� AS ��������" & _
        " From �����ʻ� A,��������Ŀ¼ B,������Ϣ C" & _
        " Where Nvl(A.����ID,0)=B.ID(+) And A.����=[1] And A.����ID=[2] And A.����ID=C.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ�뼲����Ϣ", TYPE_��Ϫũҽ, mlng����ID)
    
    With rsTemp
        Me.txt���� = !����
        Me.txt�Ա� = !�Ա�
        Me.txt������Ϣ.Tag = Nvl(!����ID, 0)
        Me.txt������Ϣ.Text = Nvl(!��������, "")
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
            txt������Ϣ.Tag = rsTemp!ID
            txt������Ϣ.Text = "(" & rsTemp!���� & ")" & rsTemp!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            If Not blnCancel Then
                MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
            End If
            If lbl������Ϣ.Tag <> "" Then txt������Ϣ.Text = lbl������Ϣ.Tag
            txt������Ϣ.SetFocus
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
