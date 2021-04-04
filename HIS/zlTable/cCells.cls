VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cCell"
Attribute VB_Ext_KEY = "Member0" ,"cCell"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlngKey As Long

Public Property Get Item(vntIndexKey As Variant) As cCell
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'################################################################################################################
'## 功能：  添加一个已经存在的对象到集合中
'##
'## 参数：  blnWithKeys :表示是否同时复制Key值
'################################################################################################################
Public Function AddExistNode(vData As cCell, Optional blnWithKeys As Boolean = False) As Long
    Dim strKey As String
    If blnWithKeys = False Then
        mlngKey = mlngKey + 1
        strKey = "K" & mlngKey
    Else
        strKey = "K" & vData.Key
    End If
    
    mCol.Add vData.Clone(blnWithKeys), strKey
    
    If blnWithKeys = False Then
        mCol(strKey).Key = mlngKey
        AddExistNode = mlngKey
    Else
        AddExistNode = vData.Key
        Call UpdateMaxKey
    End If
End Function

'################################################################################################################
'## 功能：  创建新对象
'##
'## 参数：  lngKey :用于恢复旧的集合对象，操作完后需要更新最大流水号。
'################################################################################################################
Public Function Add(Optional lngKey As Long = 0, Optional ByVal Row As Long = 0, Optional ByVal Col As Long = 0) As Long
    Dim objNewMember As cCell
    Dim sKey As String
    Set objNewMember = New cCell
    
    If Row > 0 Then objNewMember.Row = Row
    If Col > 0 Then objNewMember.Col = Col
    
    If lngKey = 0 Then
        mlngKey = mlngKey + 1
        sKey = "K" & mlngKey
        objNewMember.Key = mlngKey
        mCol.Add objNewMember, sKey
    Else
        objNewMember.Key = lngKey
        mCol.Add objNewMember, "K" & lngKey
        Call UpdateMaxKey
    End If
    
    Add = mlngKey
    Set objNewMember = Nothing
End Function

'################################################################################################################
'## 功能：  更新最大流水号
'################################################################################################################
Public Sub UpdateMaxKey()
    Dim i As Long, j As Long
    j = 0
    For i = 1 To mCol.Count
        If mCol(i).Key > j Then j = mCol(i).Key
    Next
    mlngKey = j
End Sub

'################################################################################################################
'## 功能：  复制单元格集合
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cCells
    Dim cNodes As New cCells, lngKey As Long
    Dim i As Long
    For i = 1 To mCol.Count
        lngKey = cNodes.Add(Row:=mCol(i).Row, Col:=mCol(i).Col)
        If blnWithKeys Then cNodes("K" & lngKey).Key = mCol(i).Key
        cNodes("K" & lngKey).Row = mCol(i).Row
        cNodes("K" & lngKey).Col = mCol(i).Col
        cNodes("K" & lngKey).Margin = mCol(i).Margin
        cNodes("K" & lngKey).SingleLine = mCol(i).SingleLine
        cNodes("K" & lngKey).MergeInfo = mCol(i).MergeInfo
        cNodes("K" & lngKey).Selected = mCol(i).Selected
        cNodes("K" & lngKey).Hot = mCol(i).Hot
        cNodes("K" & lngKey).Visibled = mCol(i).Visibled
        cNodes("K" & lngKey).Width = mCol(i).Width
        cNodes("K" & lngKey).Height = mCol(i).Height
        cNodes("K" & lngKey).FixedWidth = mCol(i).FixedWidth
        cNodes("K" & lngKey).AutoHeight = mCol(i).AutoHeight
        cNodes("K" & lngKey).Icon = mCol(i).Icon
        cNodes("K" & lngKey).Text = mCol(i).Text
        cNodes("K" & lngKey).Tag = mCol(i).Tag
        cNodes("K" & lngKey).FormatString = mCol(i).FormatString
        cNodes("K" & lngKey).Indent = mCol(i).Indent
        cNodes("K" & lngKey).HAlignment = mCol(i).HAlignment
        cNodes("K" & lngKey).VAlignment = mCol(i).VAlignment
        cNodes("K" & lngKey).ForeColor = mCol(i).ForeColor
        cNodes("K" & lngKey).BackColor = mCol(i).BackColor
        cNodes("K" & lngKey).GridLineColor = mCol(i).GridLineColor
        cNodes("K" & lngKey).GridLineWidth = mCol(i).GridLineWidth
        cNodes("K" & lngKey).FontName = mCol(i).FontName
        cNodes("K" & lngKey).FontSize = mCol(i).FontSize
        cNodes("K" & lngKey).FontBold = mCol(i).FontBold
        cNodes("K" & lngKey).FontItalic = mCol(i).FontItalic
        cNodes("K" & lngKey).FontStrikeout = mCol(i).FontStrikeout
        cNodes("K" & lngKey).FontUnderline = mCol(i).FontUnderline
        cNodes("K" & lngKey).FontWeight = mCol(i).FontWeight
        cNodes("K" & lngKey).Protected = mCol(i).Protected
        cNodes("K" & lngKey).ToolTipText = mCol(i).ToolTipText
        cNodes("K" & lngKey).Dirty = mCol(i).Dirty
        cNodes("K" & lngKey).Picture = mCol(i).Picture
    Next i
    Set Clone = cNodes
End Function

