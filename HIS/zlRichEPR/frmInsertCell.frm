VERSION 5.00
Begin VB.Form frmInsertCell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���뵥Ԫ��"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2790
   Icon            =   "frmInsertCell.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   330
      Left            =   1455
      TabIndex        =   5
      Top             =   1755
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   330
      Left            =   150
      TabIndex        =   4
      Top             =   1755
      Width           =   1185
   End
   Begin VB.OptionButton Option1 
      Caption         =   "����(&C)"
      Height          =   285
      Index           =   3
      Left            =   375
      TabIndex        =   3
      Top             =   1278
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "����(&R)"
      Height          =   285
      Index           =   2
      Left            =   375
      TabIndex        =   2
      Top             =   972
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "���Ԫ������(&D)"
      Height          =   285
      Index           =   1
      Left            =   375
      TabIndex        =   1
      Top             =   666
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "���Ԫ������(&I)"
      Height          =   285
      Index           =   0
      Left            =   375
      TabIndex        =   0
      Top             =   360
      Width           =   2040
   End
   Begin VB.Frame fraLine1 
      Height          =   30
      Left            =   630
      TabIndex        =   7
      Top             =   225
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   135
      Width           =   420
   End
End
Attribute VB_Name = "frmInsertCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngID As Long           '��ǰ���뷽ʽID
Private blnCancel As Boolean    '�Ƿ�ȡ���༭
Private Table As cEPRTable      '������

Public Sub ShowMe(frmParent As frmTableEditor, oTable As cEPRTable)
    '�����ӿ�
    On Error Resume Next
    Dim i As Long, CurRow As Long, CurCol As Long
    Dim lngKey As Long, lngKey2 As Long
    Dim cellFmt As F1CellFormat
        
    Set Table = oTable
    Me.Show vbModal, frmParent
    If blnCancel Then Exit Sub
    With frmParent
        Select Case lngID
        Case 0
            If .F1Book1.Visible Then
                On Error GoTo LL
                .F1Book1.InsertRange .F1Book1.Row, .F1Book1.Col, .F1Book1.Row, .F1Book1.Col, F1ShiftHorizontal
                .F1Book1.MaxCol = .F1Book1.MaxCol + 1
            End If
        Case 1
            If .F1Book1.Visible Then
                On Error GoTo LL
                .F1Book1.InsertRange .F1Book1.Row, .F1Book1.Col, .F1Book1.Row, .F1Book1.Col, F1ShiftVertical
                .F1Book1.MaxRow = .F1Book1.MaxRow + 1
            End If
        Case 2
            If .F1Book1.Visible Then
                On Error GoTo LL
                .F1Book1.InsertRange .F1Book1.Row, .F1Book1.Col, .F1Book1.Row, .F1Book1.Col, F1ShiftRows
                .F1Book1.MaxRow = .F1Book1.MaxRow + 1
            End If
        Case 3
            If .F1Book1.Visible Then
                On Error GoTo LL
                .F1Book1.InsertRange .F1Book1.Row, .F1Book1.Col, .F1Book1.Row, .F1Book1.Col, F1ShiftCols
                .F1Book1.MaxCol = .F1Book1.MaxCol + 1
            End If
        End Select
    End With
    Exit Sub
LL:
    MsgBox "����ʧ��", vbOKOnly + vbInformation, gstrSysName
End Sub

Private Sub cmdOK_Click()
    blnCancel = False
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    blnCancel = True
    Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)
    lngID = Index
End Sub
