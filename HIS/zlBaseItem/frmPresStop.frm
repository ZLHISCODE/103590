VERSION 5.00
Begin VB.Form frmPresStop 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Աͣ��"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3450
   Icon            =   "frmPresStop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   1100
   End
   Begin VB.TextBox txtStop 
      Height          =   1335
      Left            =   120
      MaxLength       =   100
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ͣ��ԭ��"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmPresStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlngPresId As Long

Public Function �༭��Ա(ByVal strID As String) As Boolean
    mlngPresId = Val(strID)
    
    frmPresStop.Show vbModal
    Exit Function
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strStop As String
    
    On Error GoTo errHandle
    
    strStop = Trim(txtStop.Text)
    
    If LenB(StrConv(strStop, vbFromUnicode)) > 100 Then
        MsgBox "ͣ��ԭ��˵��̫��(���100���ַ���50������)���������Ϊ" & LenB(StrConv(strStop, vbFromUnicode)) & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    gstrSQL = "Zl_��Ա��_ͣ��(" & mlngPresId & ",'" & strStop & "')"
            
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Unload Me
    Exit Sub
errHandle:
    If ERRCENTER() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
