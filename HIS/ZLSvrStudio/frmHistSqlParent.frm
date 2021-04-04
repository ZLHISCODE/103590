VERSION 5.00
Begin VB.Form frmHistSqlParent 
   Caption         =   "被阻塞会话历史SQL"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13605
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
   ScaleHeight     =   9270
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmHistSqlParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintType As Integer     '执行计划来源:  1-v$视图 2-历史数据
Private mdtStart As Date
Private mdtEnd As Date
Private mfrmHist As New frmHistSql
Private mlngSid As Long
Private mlngSerial As Long

Public Sub ShowMe(ByVal lngSid As Long, ByVal lngSerial As Long, ByVal dtStart As String, ByVal dtEnd As Date, ByVal intType As Integer)
    mintType = intType
    mlngSid = lngSid
    mlngSerial = lngSerial
    mdtStart = dtStart
    mdtEnd = dtEnd
    Me.Caption = "被阻塞会话(" & lngSid & "," & lngSerial & ")历史SQL"
    Me.Show
End Sub

Private Sub Form_Load()
    mfrmHist.ShowMe mlngSid, mlngSerial, mdtStart, mdtEnd, mintType
    SetParent mfrmHist.hwnd, Me.hwnd
    mfrmHist.ZOrder 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    mfrmHist.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmHist
End Sub

