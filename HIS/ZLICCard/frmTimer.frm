VERSION 5.00
Begin VB.Form frmTimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "计时器"
   ClientHeight    =   720
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   2655
   Icon            =   "frmTimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   2655
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   360
      Top             =   120
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjICCard As clsICCard
Private mblnReading As Boolean
Private Const GWL_STYLE = (-16)
Private Const WS_DISABLED = &H8000000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Sub Init(objICCard As clsICCard)
    Set mobjICCard = objICCard
End Sub

Private Sub tmrMain_Timer()
    Dim strNo As String
    If (GetWindowLong(mobjICCard.GetParent, GWL_STYLE) And WS_DISABLED) <> WS_DISABLED Then
        If mblnReading = False Then
            mblnReading = True
            strNo = mobjICCard.Read_Card
            '108227:李南春,2017/5/8，消除卡号中的空字符，是因为定长字符串引起的
            strNo = Replace(strNo, Chr(0), "")
            Call mobjICCard.ShowICCardInfo(strNo)
            mblnReading = False
        End If
    End If
End Sub

