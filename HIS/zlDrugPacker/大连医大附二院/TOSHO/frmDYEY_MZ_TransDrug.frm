VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDYEY_MZ_TransDrug 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "药品基础数据上传，请等待"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5610
   Icon            =   "frmDYEY_MZ_TransDrug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   5610
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ProgressBar prgLoadFile 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "frmDYEY_MZ_TransDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MINTDRUG As Integer = 1
Private Const MINTSTORE As Integer = 2
Private Const MINTDEPT As Integer = 3

Public Sub ChangePrg(ByVal lngCur As Long, ByVal lngSum As Long, ByVal intType As Integer)
    If lngCur = 1 Then
        If intType = MINTDRUG Then
            Me.Caption = "药品基础数据上传，请等待......"
        ElseIf intType = MINTDEPT Then
            Me.Caption = "部门基础数据上传，请等待......"
        Else
            Me.Caption = "药品库存数据上传，请等待......"
        End If
    End If
    
    Me.prgLoadFile.Value = lngCur / lngSum
    If lngCur / lngSum = 1 Then
        Unload Me
    End If
End Sub

Public Sub UnloadMe()
    Unload Me
End Sub

