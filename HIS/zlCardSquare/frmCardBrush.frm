VERSION 5.00
Begin VB.Form frmCardBrush 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "自动刷卡"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   1740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmCardBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjCard As clsBrushSequareCard
Private mblnReading As Boolean
Public Sub Init(objCard As clsBrushSequareCard)
    Set mobjCard = objCard
End Sub

Private Sub tmrMain_Timer()
    Dim strNo As String
    If mblnReading = False Then
        mblnReading = True
           If mobjCard.zlReadCard(Me, strNo) = True Then
            Call mobjCard.zlBrushCarding(strNo)
         Else
        End If
        mblnReading = False
    End If
End Sub
 
