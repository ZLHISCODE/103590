VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmPDF 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin AcroPDFLibCtl.AcroPDF AcroPDF 
      Height          =   1935
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   3255
      _cx             =   5080
      _cy             =   5080
   End
End
Attribute VB_Name = "frmPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    AcroPDF.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Function LoadFile(ByVal strFile As String) As Boolean
'功能：加载文件
    LoadFile = AcroPDF.LoadFile(strFile)
End Function

Public Function PrintFile(ByVal intType As Integer) As Boolean
'功能：打印
'参数: intType 打印方式,0-直接打印,1-交互打印
    If intType = 0 Then
        AcroPDF.printAll
    ElseIf intType = 1 Then
        AcroPDF.printWithDialog
    End If
End Function

Public Function WaitTime(ByVal lng序号 As Long, ByVal strFilePath As String, ByVal strName As String) As String
'功能:打印等待
'参数:strFilePath文件路径
'     strName 报告名称
    WaitTime = frmWait.ShowMe(Me, lng序号, strFilePath, strName)
End Function
