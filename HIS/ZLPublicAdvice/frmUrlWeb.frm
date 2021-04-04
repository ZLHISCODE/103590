VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmUrlWeb 
   Caption         =   "web´°Ìå"
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17985
   Icon            =   "frmUrlWeb.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   17985
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin SHDocVwCtl.WebBrowser webUrl 
      Height          =   6255
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   8775
      ExtentX         =   15478
      ExtentY         =   11033
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmUrlWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowMe(strUrl As String, strName As String)
    On Error Resume Next
    Me.Caption = strName
    webUrl.Navigate strUrl
    Me.Show
End Function

Private Sub Form_Resize()
    On Error Resume Next
    webUrl.Top = 0: webUrl.Left = 0
    webUrl.Width = Me.Width - 200: webUrl.Height = Me.Height - 500
End Sub

Private Sub webUrl_NewWindow2(ppDisp As Object, Cancel As Boolean)
    On Error Resume Next
    Cancel = True
    webUrl.Navigate2 webUrl.Document.activeElement.href
End Sub
