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
   StartUpPosition =   3  '����ȱʡ
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
'���ܣ������ļ�
    LoadFile = AcroPDF.LoadFile(strFile)
End Function

Public Function PrintFile(ByVal intType As Integer) As Boolean
'���ܣ���ӡ
'����: intType ��ӡ��ʽ,0-ֱ�Ӵ�ӡ,1-������ӡ
    If intType = 0 Then
        AcroPDF.printAll
    ElseIf intType = 1 Then
        AcroPDF.printWithDialog
    End If
End Function

Public Function WaitTime(ByVal lng��� As Long, ByVal strFilePath As String, ByVal strName As String) As String
'����:��ӡ�ȴ�
'����:strFilePath�ļ�·��
'     strName ��������
    WaitTime = frmWait.ShowMe(Me, lng���, strFilePath, strName)
End Function
