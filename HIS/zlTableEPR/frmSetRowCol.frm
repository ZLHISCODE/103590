VERSION 5.00
Begin VB.Form frmSetRowCol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "行列调整"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2550
   Icon            =   "frmSetRowCol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   1350
      TabIndex        =   3
      Top             =   495
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   90
      TabIndex        =   2
      Top             =   495
      Width           =   1100
   End
   Begin VB.TextBox txtValue 
      Height          =   270
      Left            =   1335
      TabIndex        =   1
      Top             =   157
      Width           =   1110
   End
   Begin VB.Label lblType 
      Caption         =   "行列(毫米)："
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   1110
   End
End
Attribute VB_Name = "frmSetRowCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mlngValue As Long '堤单位长度
Public Sub SetRowCol(ByVal frmPar As Object, ByVal strType As String, lngValue As Long)
    mlngValue = lngValue: mblnOK = False
    If strType = "行高" Then
        Me.Caption = "行高调整"
        lblType.Caption = "行高(毫米)："
        txtValue.Text = Me.ScaleY(mlngValue, vbTwips, vbMillimeters)
    Else
        Me.Caption = "列宽调整"
        lblType.Caption = "列宽(毫米)："
        txtValue.Text = Me.ScaleX(mlngValue, vbTwips, vbMillimeters)
    End If
    txtValue.Text = Format(Round(txtValue.Text, 2), "#.##")
    zlControl.TxtSelAll txtValue
    Me.Show 1, frmPar
    If mblnOK Then
        lngValue = mlngValue
    Else
        lngValue = -1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    mblnOK = True
    If Me.Caption = "行高调整" Then
        mlngValue = Int(Me.ScaleY(txtValue.Text, vbMillimeters, vbTwips))
    Else
        mlngValue = Int(Me.ScaleX(txtValue.Text, vbMillimeters, vbTwips))
    End If
    Unload Me
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdOk_Click
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
