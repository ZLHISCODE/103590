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
   NegotiateMenus  =   0   'False
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
Private mlngType As Long



Public Function ShowParameterWindow(ByVal blnIsAllInput As Boolean, owner As Object, Optional strPara As String, Optional lngType As Long) As String
'lngType: 0-所有；1-系统；2-业务
    
    ShowParameterWindow = ""
    
    Me.mblnIsAllInput = blnIsAllInput
    mstrPara = strPara
    mlngType = lngType
    
    Call SetFontSize(gbytFontSize)
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
    
    '加载系统参数
    
    If mlngType <> 2 Then
        arrPara = Split(gstrPara, ",")
        For i = 0 To UBound(arrPara)
            lstPar.AddItem arrPara(i)
        Next
    End If
    
    If mlngType <> 1 Then
        arrPara = Split(gstrBasePara, ",")
        For i = 0 To UBound(arrPara)
            lstPar.AddItem arrPara(i)
        Next
    End If
    
    If Len(mstrPara) > 0 Then
        arrPara = Split(mstrPara, "|")
        For i = 0 To UBound(arrPara)
            lstPar.AddItem "[" & arrPara(i) & "]"
        Next
    End If
End Sub

Private Sub Form_Resize()
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
    cmdSure.Left = cmdCancel.Left - cmdSure.Width - 60
End Sub

Private Sub lstPar_DblClick()
    On Error GoTo errHandle
    
    Call cmdSure_Click
    Call LoadInputPar
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub SetFontSize(ByVal bytFontSize As Byte)
    Dim lngCmdHeight As Long
    Dim lngCmdWithd As Long
    
    If bytFontSize = 9 Then
        lngCmdHeight = 350
        lngCmdWithd = 1100
    ElseIf bytFontSize = 12 Then
        lngCmdHeight = 385
        lngCmdWithd = 1300
    ElseIf bytFontSize = 15 Then
        lngCmdHeight = 420
        lngCmdWithd = 1500
    End If
    
    lstPar.FontSize = bytFontSize

    cmdCancel.FontSize = bytFontSize
    cmdCancel.Height = lngCmdHeight
    cmdCancel.Width = lngCmdWithd
    
    
    cmdSure.FontSize = bytFontSize
    cmdSure.Height = lngCmdHeight
    cmdSure.Width = lngCmdWithd
End Sub
