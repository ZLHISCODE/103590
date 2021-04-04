VERSION 5.00
Begin VB.Form frmSetPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "插入参数"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3630
   Icon            =   "frmSetPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ListBox lstPar 
      Height          =   3660
      ItemData        =   "frmSetPara.frx":6852
      Left            =   120
      List            =   "frmSetPara.frx":6854
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3885
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3885
      Width           =   1095
   End
End
Attribute VB_Name = "frmSetPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnIsOK As Boolean
Public mstrSelectPar As String
Public mblnIsAllInput As Boolean
Private mstrPara As String



Public Function ShowParameterWindow(ByVal blnIsAllInput As Boolean, owner As Object, Optional strPara As String) As String
    ShowParameterWindow = ""
    
    Me.mblnIsAllInput = blnIsAllInput
    mstrPara = strPara
    
    Call Me.Show(1, owner)
    
    If Me.mblnIsOK Then
        ShowParameterWindow = Me.mstrSelectPar
    End If
End Function

Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    mblnIsOK = False
    
    Call Me.Hide
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cmdSure_Click()
    On Error GoTo errHandle
    Dim i As Long
    
    For i = 0 To lstPar.ListCount - 1
        If lstPar.Selected(i) Then
            mstrSelectPar = lstPar.List(i)
            Exit For
        End If
    Next i
    
    mblnIsOK = True
    Call Me.Hide
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle

    mblnIsOK = False
    
    Call LoadInputPar
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub


Private Sub LoadInputPar()
'载入系统参数
    Dim i As Long
    Dim arrPara() As String
    
    arrPara = Split(gstrPara, ",")
    For i = 0 To UBound(arrPara)
        lstPar.AddItem arrPara(i)
    Next
    
    arrPara = Split(mstrPara, "|")
    For i = 0 To UBound(arrPara)
        lstPar.AddItem "[" & arrPara(i) & "]"
    Next
End Sub

Private Sub lstPar_DblClick()
    On Error GoTo errHandle
    
    Call cmdSure_Click
    Call LoadInputPar
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

