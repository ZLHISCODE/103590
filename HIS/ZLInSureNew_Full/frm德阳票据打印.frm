VERSION 5.00
Begin VB.Form frm����Ʊ�ݴ�ӡ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡ��������"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frm����Ʊ�ݴ�ӡ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   5
      Top             =   2325
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2085
      TabIndex        =   4
      Top             =   2325
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   2145
      Width           =   9300
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   2
      Top             =   750
      Width           =   7110
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   0
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1095
      Width           =   2760
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   1
      Left            =   1485
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "�ε�λ"
      Top             =   1508
      Width           =   2760
   End
   Begin VB.Image img 
      Height          =   555
      Left            =   75
      Picture         =   "frm����Ʊ�ݴ�ӡ.frx":020A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "��������Ҫ��ӡ��סԺ�ŷ�Χ"
      Height          =   165
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   420
      Width           =   4965
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ʼסԺ��"
      Height          =   180
      Index           =   1
      Left            =   510
      TabIndex        =   7
      Top             =   1162
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����סԺ��"
      Height          =   180
      Index           =   2
      Left            =   510
      TabIndex        =   6
      Tag             =   "�ε�λ"
      Top             =   1575
      Width           =   900
   End
End
Attribute VB_Name = "frm����Ʊ�ݴ�ӡ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrInfor As String
Dim mblnOK As Boolean
Dim mblnChange As Boolean
Dim mstr��ʼסԺ�� As String
Dim mstr����סԺ�� As String
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim i As Long
    Dim strInfor As String
    
    mstr��ʼסԺ�� = txtEdit(0).Text
    mstr����סԺ�� = txtEdit(1).Text
    
    mblnOK = True
    Unload Me
End Sub



Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub


Public Function ShowCard(ByRef str��ʼסԺ�� As String, ByRef str����סԺ�� As String) As Boolean

    txtEdit(0).Text = str��ʼסԺ��
    txtEdit(1).Text = str����סԺ��
    
    Me.Show 1
    str��ʼסԺ�� = mstr��ʼסԺ��
    str����סԺ�� = mstr����סԺ��
    ShowCard = mblnOK

End Function

