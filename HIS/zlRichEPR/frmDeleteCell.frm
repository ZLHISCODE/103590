VERSION 5.00
Begin VB.Form frmDeleteCell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "删除单元格"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2715
   Icon            =   "frmDeleteCell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   330
      Left            =   1440
      TabIndex        =   7
      Top             =   1755
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   330
      Left            =   135
      TabIndex        =   6
      Top             =   1755
      Width           =   1185
   End
   Begin VB.OptionButton Option1 
      Caption         =   "整列(&C)"
      Height          =   285
      Index           =   3
      Left            =   270
      TabIndex        =   5
      Top             =   1260
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "整行(&R)"
      Height          =   285
      Index           =   2
      Left            =   270
      TabIndex        =   4
      Top             =   960
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "下方单元格上移(&U)"
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   3
      Top             =   660
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "右侧单元格左移(&L)"
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   2
      Top             =   360
      Width           =   2040
   End
   Begin VB.Frame fraLine1 
      Height          =   30
      Left            =   630
      TabIndex        =   1
      Top             =   225
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "删除"
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   420
   End
End
Attribute VB_Name = "frmDeleteCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngID As Long           '当前删除方式ID
Private blnCancel As Boolean    '是否取消编辑

Public Sub ShowMe(frmParent As Object)
'公共接口
    Me.Show vbModal, frmParent
    If blnCancel Then Exit Sub
    With frmParent
        Select Case lngID
        Case 0
            If .F1Book1.Visible Then
                .F1Book1.MaxCol = .F1Book1.MaxCol - 1
                .F1Book1.DeleteRange .F1Book1.Row, .F1Book1.Col, .F1Book1.Row, .F1Book1.Col, F1ShiftHorizontal
            End If
        Case 1
            If .F1Book1.Visible Then
                .F1Book1.MaxRow = .F1Book1.MaxRow - 1
                .F1Book1.DeleteRange .F1Book1.Row, .F1Book1.Col, .F1Book1.Row, .F1Book1.Col, F1ShiftVertical
            End If
        Case 2
            If .F1Book1.Visible Then
                .F1Book1.MaxRow = .F1Book1.MaxRow - 1
                .F1Book1.DeleteRange .F1Book1.Row, .F1Book1.Col, .F1Book1.Row, .F1Book1.Col, F1ShiftRows
            End If
        Case 3
            If .F1Book1.Visible Then
                .F1Book1.MaxCol = .F1Book1.MaxCol - 1
                .F1Book1.DeleteRange .F1Book1.Row, .F1Book1.Col, .F1Book1.Row, .F1Book1.Col, F1ShiftCols
            End If
        End Select
'        .cTable.高度 = .F1Book1.Height
'        .cTable.宽度 = .F1Book1.Width
'        .cTable.行数 = .F1Book1.MaxRow
'        .cTable.列数 = .F1Book1.MaxCol
    End With
End Sub

Private Sub Command1_Click()
    blnCancel = False
    Unload Me
End Sub

Private Sub Command2_Click()
    blnCancel = True
    Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)
    lngID = Index
End Sub
