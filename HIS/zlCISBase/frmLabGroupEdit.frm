VERSION 5.00
Begin VB.Form frmLabGroupEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����С��"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "frmLabGroupEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5145
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chk���� 
      Caption         =   "��������"
      Height          =   180
      Left            =   450
      TabIndex        =   6
      Top             =   1275
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3405
      TabIndex        =   5
      Top             =   1215
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2145
      TabIndex        =   4
      Top             =   1215
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1155
      TabIndex        =   3
      Top             =   555
      Width           =   3570
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1155
      TabIndex        =   1
      Top             =   195
      Width           =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "С������"
      Height          =   270
      Left            =   315
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "С�����"
      Height          =   270
      Left            =   285
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmLabGroupEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngGroupID As Long
Private mblnAdd As Boolean
Private mblnOK As Boolean
Private mstr�������� As String

Public Function ShowMe(ByRef lngItemID As Long, ByVal intAdd As Integer, ByVal str�������� As String, ByVal frmMain As Form) As Boolean
    mlngGroupID = lngItemID
    mblnAdd = intAdd = 1
    mstr�������� = str��������
    mblnOK = False
    Me.Show vbModal, frmMain
    ShowMe = mblnOK
    lngItemID = mlngGroupID
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, str���� As String, str���� As String
    
    On Error GoTo ErrHandle
    str���� = Replace(Trim(txt����.Text), "'", "")
    str���� = Replace(Trim(txt����.Text), "'", "")
    If str���� = "" Or str���� = "" Then
        MsgBox "��������Ʋ���Ϊ�գ�", vbInformation, Me.Caption
        Exit Sub
    End If
    If txt����.Tag = "����" Then
        mlngGroupID = zlDatabase.GetNextId("���ű�")
        strSQL = "zl_����С��_Edit(1," & mlngGroupID & ",'" & str���� & "','" & str���� & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Else
        strSQL = "zl_����С��_Edit(2," & mlngGroupID & ",'" & str���� & "','" & str���� & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    mblnOK = True
    If chk����.Value = 1 Then
        txt����.Text = ""
        txt����.Text = ""
    Else
        Unload Me
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    If mblnAdd Then
        txt����.Tag = "����"
        txt����.Text = ""
        txt����.Text = ""
    Else
        txt����.Tag = "�༭"
        txt����.Text = Split(mstr��������, "|")(0)
        txt����.Text = Split(mstr��������, "|")(1)
    End If
End Sub
