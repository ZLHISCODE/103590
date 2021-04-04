VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmAwrSql 
   Caption         =   "Form1"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAwrSql.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser webAwr 
      Height          =   4695
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   4935
      ExtentX         =   8705
      ExtentY         =   8281
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
Attribute VB_Name = "frmAwrSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFile As String

Public Sub ShowMe(strSqlID As String, strFilePath As String)
    Me.Caption = " 语句" & strSqlID & " AWR报告"
    mstrFile = strFilePath
    webAwr.Navigate "file:///" & Replace(strFilePath, "\", "/")
    Me.Show 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    webAwr.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '退出窗体时,删除临时文件
    gobjFile.DeleteFile mstrFile
End Sub
