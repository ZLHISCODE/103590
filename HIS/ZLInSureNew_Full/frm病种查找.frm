VERSION 5.00
Begin VB.Form frm���ֲ��� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frm���ֲ���.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2220
      TabIndex        =   4
      Top             =   900
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   960
      TabIndex        =   3
      Top             =   900
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -30
      TabIndex        =   2
      Top             =   690
      Width           =   5085
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   870
      TabIndex        =   1
      Top             =   210
      Width           =   2655
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   270
      Width           =   360
   End
End
Attribute VB_Name = "frm���ֲ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lng���� As Long
Private rsFind As New ADODB.Recordset

Private Sub cmdFind_Click()
    Dim strInput As String
    strInput = Trim(UCase(txt����.Text))
    
    If Val(cmdFind.Tag) = 0 Then
        gstrSQL = " Select ID From ���ղ��� " & _
                  " Where (���� Like '%" & strInput & "%' Or ���� Like '%" & strInput & "%') And ����=" & lng����
        Call OpenRecordset(rsFind, "��ȡ���Ҽ�¼��")
        If rsFind.RecordCount = 0 Then
            MsgBox "û���ҵ��κμ�¼�������䣡", vbInformation, gstrSysName
            txt����.SetFocus
            Exit Sub
        Else
            cmdFind.Tag = 1
            cmdFind.Caption = "��һ��(&N)"
        End If
    Else
        rsFind.MoveNext
        If rsFind.EOF Then
            If MsgBox("����Ҫ��ͷ��ʼ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                txt����.SetFocus
                cmdFind.Tag = 0
                cmdFind.Caption = "����(&F)"
                Exit Sub
            Else
                rsFind.MoveFirst
            End If
        End If
    End If
    
    '���ݼ�¼����λ������
    With frm���ղ���
        Dim lngRow As Long, lngItems As Long
        lngItems = .lvwItem.ListItems.Count
        For lngRow = 1 To lngItems
            If Val(Mid(.lvwItem.ListItems(lngRow).Key, 2)) = rsFind!ID Then
                .lvwItem.ListItems(lngRow).Selected = True
                .lvwItem.SelectedItem.Selected = True
                .lvwItem.SelectedItem.EnsureVisible
                Exit For
            End If
        Next
    End With
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub
