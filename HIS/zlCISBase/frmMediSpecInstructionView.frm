VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMediSpecInstructionView 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "使用说明预览"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8385
   Icon            =   "frmMediSpecInstructionView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8385
   StartUpPosition =   1  '所有者中心
   Begin RichTextLib.RichTextBox rtbDetails 
      Height          =   5430
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   9578
      _Version        =   393217
      BackColor       =   14737632
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMediSpecInstructionView.frx":000C
   End
End
Attribute VB_Name = "frmMediSpecInstructionView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowMe(ByVal frmParent As Object, ByVal frmName As String)
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Me.Caption = "【" & frmName & "】" & "使用说明预览"
    Me.Show vbModal, frmParent
End Sub

