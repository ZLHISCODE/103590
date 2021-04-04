VERSION 5.00
Begin VB.Form frmUnloadWebKixt 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "frmUnloadWebKixt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CloseWebKitX()
    On Error Resume Next
    If Err <> 0 Then Err.Clear
End Sub
