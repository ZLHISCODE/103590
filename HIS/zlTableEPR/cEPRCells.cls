VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cEPRCell"
Attribute VB_Ext_KEY = "Member0" ,"cEPRCell"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlngIndex As Long
Private mlngRows As Long
Private mlngCols As Long
Public Property Get Rows() As Long
    Rows = mlngRows
End Property
Public Property Let Rows(ByVal vData As Long)
    mlngRows = vData
End Property
Public Property Get Cols() As Long
    Cols = mlngCols
End Property
Public Property Let Cols(ByVal vData As Long)
    mlngCols = vData
End Property
Public Property Get Item(vntIndexKey As Variant) As cTabCell
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property
Public Property Get Cell(ByVal lngRow As Long, ByVal lngCol As Long) As cTabCell
    On Error GoTo FindCell
    Set Cell = Me.Item("K" & lngRow & "_" & lngCol)
    Exit Sub
'先用Key方式找，找不到才用循环
FindCell:
Err.Clear
Dim l As Long, blnDesc As Boolean
    blnDesc = (lngRow > Me.Rows / 2 Or lngCol > Me.Cols / 2)
    For l = IIf(blnDesc, Me.Count, 1) To IIf(blnDesc, 1, Me.Count) Step IIf(blnDesc, -1, 1)
        With Me.Item(l)
            If .Row = lngRow And .Col = lngCol Then
                Set Cell = Me.Item(.Key)
                Exit For
            End If
        End With
    Next
End Property
Public Property Get OwnCell(ByVal lngRow As Long, ByVal lngCol As Long) As cTabCell
'功能:传入的单元格被哪个单元格合并,返回首单元格,不处理传入的单元格为非合并的情况
Dim l As Long, blnDesc As Boolean
Dim lsRow As Long, leRow As Long, lsCol As Long, leCol As Long
    blnDesc = (lngRow > Me.Rows / 2 Or lngCol > Me.Cols / 2)
    For l = IIf(blnDesc, Me.Count, 1) To IIf(blnDesc, 1, Me.Count) Step IIf(blnDesc, -1, 1)
        With Me.Item(l)
            If .Merge And InStr(.MergeRange, ";") > 0 Then '只查有效的合并单元格
                lsRow = Split(Split(.MergeRange, ";")(0), ",")(0): leRow = Split(Split(.MergeRange, ";")(1), ",")(0)
                lsCol = Split(Split(.MergeRange, ";")(0), ",")(1): leCol = Split(Split(.MergeRange, ";")(1), ",")(1)
                
                If lngRow >= lsRow And lngRow <= leRow And lngCol >= lsCol And lngCol <= leCol Then
                    Set OwnCell = Me.Item(.Key)
                    Exit For
                End If
            End If
        End With
    Next
End Property
Public Property Get Count() As Long
    Count = mCol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub
Public Sub RemoveAll()
    Set mCol = Nothing
    Set mCol = New Collection
    mlngIndex = 0
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'################################################################################################################
'## 功能：  创建新对象
'##
'## 参数：  lngKey :用于恢复旧的集合对象，操作完后需要更新最大流水号。
'################################################################################################################
Public Function Add(ByVal Row As Long, ByVal Col As Long) As String
    Dim objNewMember As cTabCell
    Dim sKey As String
    Set objNewMember = New cTabCell
    
    If Row > 0 Then objNewMember.Row = Row
    If Col > 0 Then objNewMember.Col = Col
    
    mlngIndex = mlngIndex + 1: sKey = "K" & Row & "_" & Col
    objNewMember.Index = mlngIndex: objNewMember.Key = sKey
    mCol.Add objNewMember, sKey
    
    Add = sKey
    Set objNewMember = Nothing
End Function
