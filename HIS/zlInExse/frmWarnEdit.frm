VERSION 5.00
Begin VB.Form frmWarnEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ʱ�������"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "frmWarnEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4560
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -180
      TabIndex        =   6
      Top             =   1545
      Width           =   5115
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2550
      TabIndex        =   5
      Top             =   1740
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1290
      TabIndex        =   4
      Top             =   1740
      Width           =   1100
   End
   Begin VB.ComboBox cboCopy 
      Height          =   300
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   930
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   1
      Top             =   435
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Left            =   615
      TabIndex        =   2
      Top             =   990
      Width           =   720
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   615
      TabIndex        =   0
      Top             =   495
      Width           =   720
   End
End
Attribute VB_Name = "frmWarnEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrCopy As String
Private mstrName As String

Public Function ShowMe(frmParent As Object, ByVal strShcemes As String, strCopy As String) As String
    Dim i As Long
    
    Load Me
    
    Me.cboCopy.AddItem ""
    cboCopy.ListIndex = 0
    For i = 0 To UBound(Split(strShcemes, ","))
        cboCopy.AddItem Split(strShcemes, ",")(i)
    Next
    
    strCopy = ""
    Me.Show 1, frmParent
    
    strCopy = mstrCopy
    ShowMe = mstrName
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    If txtName.Text = "" Then
        MsgBox "�����뷽�����ơ�", vbInformation, gstrSysName
        txtName.SetFocus: Exit Sub
    End If
    
    If zlCommFun.ActualLen(txtName.Text) > txtName.MaxLength Then
        MsgBox "�������ƹ������������ " & txtName.MaxLength \ 2 & " �����ֻ� " & txtName.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txtName.SetFocus: Exit Sub
    End If
    
    For i = 0 To cboCopy.ListCount - 1
        If cboCopy.List(i) = txtName.Text Then
            MsgBox "�������Ѿ����ڣ������������������ơ�", vbInformation, gstrSysName
            txtName.SetFocus: Exit Sub
        End If
    Next
    
    mstrName = txtName.Text
    mstrCopy = cboCopy.Text
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',;", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    mstrName = ""
    mstrCopy = ""
End Sub

Private Sub txtName_GotFocus()
    Call zlControl.TxtSelAll(txtName)
End Sub
