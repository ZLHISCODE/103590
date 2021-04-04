VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9645
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   15
      Top             =   945
   End
   Begin VB.TextBox Text1 
      Height          =   5340
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   165
      Width           =   9120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
Unload Me
End Sub
