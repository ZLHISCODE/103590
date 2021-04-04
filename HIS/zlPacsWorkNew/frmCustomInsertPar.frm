VERSION 5.00
Begin VB.Form frmCustomInsertPar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomInsertPar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   375
      Left            =   1125
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.ListBox lstPar 
      Height          =   3570
      ItemData        =   "frmCustomInsertPar.frx":000C
      Left            =   120
      List            =   "frmCustomInsertPar.frx":000E
      TabIndex        =   0
      Top             =   195
      Width           =   3270
   End
End
Attribute VB_Name = "frmCustomInsertPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mobjInputList As ucFlexGrid
Public mblnIsOk As Boolean
Public mstrSelectPar As String
Public mblnIsAllInput As Boolean

Public Function ShowParameterWindow(objInputList As Object, ByVal blnIsAllInput As Boolean, owner As Object) As String
    ShowParameterWindow = ""
    Set Me.mobjInputList = objInputList
    
    Me.mblnIsAllInput = blnIsAllInput
    
    Call Me.Show(1, owner)
    
    If Me.mblnIsOk Then
        ShowParameterWindow = Me.mstrSelectPar
    End If
End Function

Private Sub cmdCancel_Click()
On Error GoTo errHandle
    mblnIsOk = False
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    Dim i As Long
    
    For i = 0 To lstPar.ListCount - 1
        If lstPar.Selected(i) Then
            mstrSelectPar = lstPar.list(i)
            Exit For
        End If
    Next i
    
    mblnIsOk = True
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    mblnIsOk = False
    
    Call LoadInputPar
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadInputPar()
'�����ѡ���¼�����
    Dim i As Long
    
    lstPar.AddItem "[��ǰ����]"
    lstPar.AddItem "[��ǰʱ��]"
    lstPar.AddItem "[��ǰ�û�ID]"
    lstPar.AddItem "[��ǰ����ID]"
    lstPar.AddItem "[��ǰϵͳ���]"
    lstPar.AddItem "[��ǰģ����]"
    
    For i = 1 To IIf(mblnIsAllInput, mobjInputList.GridRows - 1, mobjInputList.SelectionRow - 1)
        If mobjInputList.Text(i, "¼����Ŀ") <> "" Then
            lstPar.AddItem "[" & mobjInputList.Text(i, "¼����Ŀ") & "]"
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandle
    Call SaveWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub lstPar_DblClick()
    Call cmdSure_Click
End Sub
