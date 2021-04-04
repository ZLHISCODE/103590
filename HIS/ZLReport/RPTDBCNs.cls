VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTDBCNs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolDBCNs As Collection

Public Function Add(ByVal NO As Integer, ByVal Name As String, ByVal Conn As ADODB.Connection) As RPTDBCN
    If NO < 0 Or Trim(Name) = "" Or Conn Is Nothing Then
        MsgBox "参数未赋值，拒绝添加对象！", vbInformation, App.Title
        Exit Function
    End If
    
    Dim clsRPTDBCN As New RPTDBCN
    
    With clsRPTDBCN
        .NO = NO
        .Name = Name
        Set .Connection = Conn
    End With
    
    mcolDBCNs.Add clsRPTDBCN, "_" & NO
    
    Set Add = clsRPTDBCN
End Function

Public Sub Remove(ByVal Index As Integer)
    On Error Resume Next
    mcolDBCNs.Remove Index
End Sub

Public Sub Clear()
    Dim i As Integer
    
    For i = mcolDBCNs.Count To 1 Step -1
        mcolDBCNs.Remove i
    Next
    Set mcolDBCNs = New Collection
End Sub

Public Property Get Item(ByVal IndexKey As Variant) As RPTDBCN
    On Error GoTo hErr
    Set Item = mcolDBCNs(IndexKey)
    Exit Property
    
hErr:
    Set Item = Nothing
End Property

Public Property Get Count() As Integer
    Count = mcolDBCNs.Count
End Property

Private Sub Class_Initialize()
    Set mcolDBCNs = New Collection
End Sub

Private Sub Class_Terminate()
    Me.Clear
    Set mcolDBCNs = Nothing
End Sub
