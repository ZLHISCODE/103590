VERSION 5.00
Begin VB.Form frmHF 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҳü��ҳ��"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmHF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4920
      TabIndex        =   1
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4920
      TabIndex        =   0
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "����"
      Height          =   615
      Left            =   3180
      Picture         =   "frmHF.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmdҳ�� 
      Caption         =   "��ҳ��"
      Height          =   615
      Left            =   1170
      Picture         =   "frmHF.frx":06F6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmdMan 
      Caption         =   "�û���"
      Height          =   615
      Left            =   4200
      Picture         =   "frmHF.frx":0DE0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "ʱ��"
      Height          =   615
      Left            =   2190
      Picture         =   "frmHF.frx":14CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmdUnit 
      Caption         =   "��λ��"
      Height          =   615
      Left            =   5220
      Picture         =   "frmHF.frx":1BB4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmdҳ�� 
      Caption         =   "ҳ��"
      Height          =   615
      Left            =   150
      Picture         =   "frmHF.frx":229E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1350
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   1365
      Index           =   3
      Left            =   4230
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2340
      Width           =   1785
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   1365
      Index           =   2
      Left            =   2250
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2340
      Width           =   1785
   End
   Begin VB.TextBox Text1 
      Height          =   1365
      Index           =   1
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2340
      Width           =   1785
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHF.frx":2988
      Height          =   705
      Left            =   270
      TabIndex        =   14
      Top             =   150
      Width           =   4125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��(&R):"
      Height          =   180
      Left            =   4230
      TabIndex        =   12
      Top             =   2070
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��(&M):"
      Height          =   180
      Left            =   2280
      TabIndex        =   10
      Top             =   2070
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��(&L):"
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   2070
      Width           =   540
   End
End
Attribute VB_Name = "frmHF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��������������ҳü��ҳ��




Dim mstrTemp As String      '��ʱ��ҳüҳ��ֵ
Dim mblnTemp As Boolean     'Ϊ�ٱ�ʾ�ǰ�"ȡ��"�رմ���
Dim mintIndex As Integer    '��ý����Text1������ֵ

Private Sub Form_Load()
    Dim intPos As Integer
    Dim intPos1 As Integer
    mblnTemp = False
    On Error Resume Next
    intPos = InStr(mstrTemp, ";")
    intPos1 = intPos + 1
    Text1(1).Text = Mid(mstrTemp, 1, intPos - 1)
    intPos = InStr(intPos1, mstrTemp, ";")
    Text1(2).Text = Mid(mstrTemp, intPos1, intPos - intPos1)
    intPos1 = intPos + 1
    Text1(3).Text = Mid(mstrTemp, intPos1)
    mintIndex = 1
    'On Error GoTo 0
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    mintIndex = Index
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrTemp = Text1(1).Text & ";" & Text1(2).Text & ";" & Text1(3).Text
    mblnTemp = True
    Unload Me
End Sub

Public Function GetText(strGet As String) As Boolean
    mstrTemp = strGet
    Me.Show 1
    strGet = mstrTemp
    GetText = mblnTemp
End Function

Private Sub cmdҳ��_Click()
    Text1(mintIndex).SelText = "��[ҳ��]ҳ"
End Sub

Private Sub cmdҳ��_Click()
    Text1(mintIndex).SelText = "��[ҳ��]ҳ"
End Sub

Private Sub cmdTime_Click()
    Text1(mintIndex).SelText = "[ʱ��]"
End Sub

Private Sub cmdDate_Click()
    Text1(mintIndex).SelText = "[����]"
End Sub

Private Sub cmdMan_Click()
    Text1(mintIndex).SelText = "[�û���]"
End Sub

Private Sub cmdUnit_Click()
    Text1(mintIndex).SelText = "[��λ��]"
End Sub
