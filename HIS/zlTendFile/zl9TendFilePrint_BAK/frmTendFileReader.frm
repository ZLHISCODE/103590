VERSION 5.00
Begin VB.Form frmTendFileReader 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "护理文件打印数据阅览器"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8625
   Icon            =   "frmTendFileReader.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin zl9TendFilePrint.usrTendFileReader usrTendFileReader 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   8599
   End
End
Attribute VB_Name = "frmTendFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnReady As Boolean          'TRUE表示已准备好打印,FALSE表示无数据无法打印

Private Sub Form_Load()
    Dim int页码 As Integer
    
    int页码 = frmAsk.intPage        '零表示续打，指定页表示重打
    blnReady = usrTendFileReader.ShowMe(Me, glng文件ID, glng病人ID, glng主页ID, gint婴儿, int页码)
    Me.Hide
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With usrTendFileReader
        .Left = 0
        .Top = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub

Public Function NextPage() As Boolean
    If usrTendFileReader.isEndPage Then Exit Function
    NextPage = usrTendFileReader.NextPage
End Function

Public Function GetPages() As Integer
    GetPages = usrTendFileReader.GetPages
End Function

Public Function GetStartPage() As Integer
    GetStartPage = usrTendFileReader.GetStartPage
End Function

Public Sub ShowPage(ByVal intPage As Integer)
    Call usrTendFileReader.ShowPage(intPage)
End Sub

Public Function PrintPage() As Boolean
    '打印当前页面数据并记录打印结果
    PrintPage = usrTendFileReader.PrintPage
End Function

Public Function GetCollectCols() As String
    GetCollectCols = usrTendFileReader.GetCollectCols
End Function

Public Function PrintHead() As Boolean
    PrintHead = usrTendFileReader.PrintHead
End Function

Public Function PrintFoot() As Boolean
    PrintFoot = usrTendFileReader.PrintFoot
End Function
