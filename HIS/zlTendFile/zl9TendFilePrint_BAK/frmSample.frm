VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   2220
      ScaleHeight     =   1575
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   780
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   555
      Left            =   1020
      TabIndex        =   0
      Top             =   1200
      Width           =   1365
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'本窗体用于得到的指定字体与大小的字符的高度和宽度
