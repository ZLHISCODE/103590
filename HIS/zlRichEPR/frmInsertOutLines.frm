VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.51#0"; "Codejock.CommandBars.Unicode.9510.ocx"
Begin VB.Form frmInsertOutLine 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1440
      ScaleHeight     =   240
      ScaleWidth      =   3525
      TabIndex        =   10
      Top             =   1485
      Width           =   3525
      Begin VB.OptionButton optNeeded 
         Caption         =   "��(&Y)"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optNeeded 
         Caption         =   "��(&N)"
         Height          =   240
         Index           =   1
         Left            =   990
         TabIndex        =   11
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1035
      Width           =   3705
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   2940
      TabIndex        =   7
      Top             =   2070
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2070
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   180
      TabIndex        =   5
      Top             =   1845
      Width           =   5010
   End
   Begin VB.TextBox txtText 
      Height          =   300
      Left            =   1395
      TabIndex        =   3
      Text            =   "δ�������"
      Top             =   630
      Width           =   3705
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1395
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "δ�������"
      Top             =   225
      Width           =   3705
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   0
      Top             =   0
      _Version        =   589875
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmInsertOutLines.frx":0000
   End
   Begin VB.Label Label4 
      Caption         =   "�Ƿ����(&I)"
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   9
      Top             =   1485
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   "��ٲ��(&L)"
      Height          =   240
      Left            =   225
      TabIndex        =   4
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "��ʾ����(&T)"
      Height          =   240
      Left            =   225
      TabIndex        =   2
      Top             =   675
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "�������(&N)"
      Height          =   240
      Left            =   225
      TabIndex        =   0
      Top             =   270
      Width           =   1140
   End
End
Attribute VB_Name = "frmInsertOutLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngID As Long
Dim ModType As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If ModType = "Add" Then
        gfrm������.InsertOutline gfrm������.Editor1.Selection.StartPos, gfrm������.Editor1.Selection.EndPos, txtName, txtText, cmbLevel.ListIndex + 1, optNeeded(0).Value
    Else
        gfrm������.ModifyOutline lngID, txtName, txtText, cmbLevel.ListIndex + 1, optNeeded(0).Value
    End If
    Unload Me
End Sub

Private Sub txtName_Change()
    If Trim(txtName) = "" Then
        txtText = "δ�������"
        txtName = txtText
    Else
        txtText = txtName
    End If
End Sub

Public Sub ModifyOutline(frmParent As Object, lID As Long)
    ModType = "Mod"
    lngID = lID
    With cmbLevel
        .Clear
        .AddItem "���һ"
        .AddItem "��ٶ�"
        .AddItem "�����"
        .AddItem "�����"
        .AddItem "�����"
        .AddItem "�����"
        .AddItem "�����"
        .AddItem "��ٰ�"
        .AddItem "��پ�"
        .ListIndex = 0
    End With
    With gfrm������
        txtName = .Document.OutLines("K" & lngID).����
        txtText = .Document.OutLines("K" & lngID).�ı�
        cmbLevel.ListIndex = .Document.OutLines("K" & lngID).��� - 1
        optNeeded(0).Value = .Document.OutLines("K" & lngID).����
        optNeeded(1).Value = Not optNeeded(0).Value
    End With
    Me.Show , frmParent
End Sub

Public Sub InsertOutline(frmParent As Object)
    ModType = "Add"
    lngID = 0
    With cmbLevel
        .Clear
        .AddItem "���һ"
        .AddItem "��ٶ�"
        .AddItem "�����"
        .AddItem "�����"
        .AddItem "�����"
        .AddItem "�����"
        .AddItem "�����"
        .AddItem "��ٰ�"
        .AddItem "��پ�"
        .ListIndex = 0
    End With
    If gfrm������.Editor1.SelLength > 0 Then
        txtText.Text = Left(gfrm������.Editor1.SelText, 50)
        txtName = txtText
    Else
        txtText = "δ�������"
        txtName = txtText
    End If
    Me.Show , frmParent
End Sub
