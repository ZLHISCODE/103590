VERSION 5.00
Begin VB.Form frmHistSqlParent 
   Caption         =   "�����Ự��ʷSQL"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHistSqlParent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "frmHistSqlParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmHist As New frmHistSql
Private mlngSid As Long
Private mlngSerial As Long

Public Sub ShowMe(ByVal lngSid As Long, ByVal lngSerial As Long)
    mlngSid = lngSid
    mlngSerial = mlngSerial
    Me.Show
End Sub

Private Sub Form_Load()
    mfrmHist.SetSid mlngSid, mlngSerial
    mfrmHist.ShowMe
    SetParent mfrmHist.hwnd, Me.hwnd
    mfrmHist.ZOrder 0
End Sub

Private Sub FindWindowAndSetActive(ByVal FrmObj As Form)
    Dim LngTargetHdl As Long
    '--����ô����Ѿ���,�򼤻���(����,����Ĵ�С���ᷢ���仯)--zyb
    LngTargetHdl = FindWindow(vbNullString, FrmObj.Caption)
    If LngTargetHdl <> 0 Then
        If IsIconic(LngTargetHdl) Then
            Call ShowWindow(LngTargetHdl, 9)            '��ԭָ������Ϊԭ��С
        End If
        Call SetActiveWindow(LngTargetHdl)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    mfrmHist.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmHist = Nothing
End Sub

