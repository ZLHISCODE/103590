VERSION 5.00
Begin VB.Form frmMipModule 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer tmrRun 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMipModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytLinkType As Byte
Private mstrLinkPara As String
Private mblnUsing As Boolean

Public Event OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)

Public Property Let Using(ByVal blnData As Boolean)
    mblnUsing = blnData
End Property

Public Property Get Using() As Boolean
    Using = mblnUsing
End Property

Public Sub OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)
    
    mbytLinkType = bytLinkType
    mstrLinkPara = strLinkPara
    tmrRun.Enabled = True
    
End Sub

Private Sub Form_Load()
    mblnUsing = False
End Sub

Private Sub tmrRun_Timer()
    
    tmrRun.Enabled = False

    RaiseEvent OpenLink(mbytLinkType, mstrLinkPara)

End Sub


