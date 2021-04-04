VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDFViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfrmPDF As New frmPDF

Public Function GetFrm() As Object
'功能：返回窗体
    Set GetFrm = mfrmPDF
End Function

Public Function LoadFile(ByVal strFile As String) As Boolean
'功能：加载文件
    LoadFile = mfrmPDF.LoadFile(strFile)
End Function

Public Function PrintFile(ByVal intType As Integer) As Boolean
'功能：打印
'参数: intType 打印方式,0-直接打印,1-交互打印
    PrintFile = mfrmPDF.PrintFile(intType)
End Function

Public Function WaitTime(ByVal lng序号 As Long, ByVal strFilePath As String, ByVal strName As String)
'功能:打印等待
'参数:strFilePath文件路径
'     strName 报告名称
    WaitTime = mfrmPDF.WaitTime(lng序号, strFilePath, strName)
End Function
