VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cEPRTable"
Attribute VB_Ext_KEY = "Member0" ,"cEPRTable"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlngKey As Long

Public Property Get Item(vntIndexKey As Variant) As cEPRTable
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
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'################################################################################################################
'## 功能：  检查所有对象的ID是否有效
'##
'## 参数：  eEditType :当前的编辑模式
'##
'## 说明：  如果没有ID，则到数据库中提取一个唯一ID号。
'################################################################################################################
Public Sub CheckValidIDs(eEditType As EditTypeEnum)
    Dim i As Long
    Select Case eEditType
    Case cprET_病历文件定义
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = zlDatabase.GetNextId("病历文件结构")
        Next
    Case cprET_全文示范编辑
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = zlDatabase.GetNextId("病历范文内容")
        Next
    Case cprET_单病历编辑, cprET_单病历审核
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = GetEPRContentNextId
        Next
    End Select
End Sub

'################################################################################################################
'## 功能：  添加一个已经存在的图片对象到集合中
'##
'## 参数：  blnWithKeys :表示是否同时复制Key值
'################################################################################################################
Public Function AddExistNode(vData As cEPRTable, Optional blnWithKeys As Boolean = False) As Long
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
Public Function Add(Optional lngKey As Long = 0) As Long
    Dim objNewMember As cEPRTable
    Set objNewMember = New cEPRTable
    Dim sKey As String

    If lngKey = 0 Then
        mlngKey = mlngKey + 1
        sKey = "K" & mlngKey
        objNewMember.Key = mlngKey
        mCol.Add objNewMember, sKey
        Add = mlngKey
    Else
        objNewMember.Key = lngKey
        mCol.Add objNewMember, "K" & lngKey
        Call UpdateMaxKey
        Add = lngKey
    End If
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
'## 功能：  复制表格集合
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cEPRTables
    '复制对象
    Dim cNodes As New cEPRTables, lngKey As Long
    Dim i As Long
    For i = 1 To mCol.Count
        lngKey = cNodes.Add()
        If blnWithKeys Then cNodes("K" & lngKey).Key = mCol(i).Key
        cNodes("K" & lngKey).ID = mCol(i).ID
        cNodes("K" & lngKey).文件ID = mCol(i).文件ID
        cNodes("K" & lngKey).开始版 = mCol(i).开始版
        cNodes("K" & lngKey).终止版 = mCol(i).终止版
        cNodes("K" & lngKey).父ID = mCol(i).父ID
        cNodes("K" & lngKey).对象序号 = mCol(i).对象序号
        cNodes("K" & lngKey).保留对象 = mCol(i).保留对象
        cNodes("K" & lngKey).是否换行 = mCol(i).是否换行
        cNodes("K" & lngKey).预制提纲ID = mCol(i).预制提纲ID
        
        cNodes("K" & lngKey).Rows = mCol(i).Rows
        cNodes("K" & lngKey).Cols = mCol(i).Cols
        cNodes("K" & lngKey).Width = mCol(i).Width
        cNodes("K" & lngKey).Height = mCol(i).Height
        cNodes("K" & lngKey).ColWidthString = mCol(i).ColWidthString
        cNodes("K" & lngKey).SingleLine = mCol(i).SingleLine
        cNodes("K" & lngKey).AlternateRowBackColor = mCol(i).AlternateRowBackColor
        cNodes("K" & lngKey).BackColor = mCol(i).BackColor
        cNodes("K" & lngKey).GridLineColor = mCol(i).GridLineColor
        cNodes("K" & lngKey).GridLineWidth = mCol(i).GridLineWidth
        cNodes("K" & lngKey).BorderColor = mCol(i).BorderColor
        cNodes("K" & lngKey).BorderWidth = mCol(i).BorderWidth
        cNodes("K" & lngKey).ForeColor = mCol(i).ForeColor
        cNodes("K" & lngKey).FontQuality = mCol(i).FontQuality
        cNodes("K" & lngKey).AutoHeight = mCol(i).AutoHeight
        cNodes("K" & lngKey).WordEllipsis = mCol(i).WordEllipsis
        cNodes("K" & lngKey).CellMargin = mCol(i).CellMargin
        cNodes("K" & lngKey).CellIndent = mCol(i).CellIndent
        cNodes("K" & lngKey).TableType = mCol(i).TableType
        cNodes("K" & lngKey).ExtendTag = mCol(i).ExtendTag
        cNodes("K" & lngKey).标记 = mCol(i).标记
        
        Set cNodes("K" & lngKey).Pictures = mCol(i).Pictures.Clone
        Set cNodes("K" & lngKey).Cells = mCol(i).Cells.Clone
        Set cNodes("K" & lngKey).Elements = mCol(i).Elements.Clone(True)
        Set cNodes("K" & lngKey).DicSpecSQLs = mCol(i).DicSpecSQLs
    Next i
    If blnWithKeys Then Call UpdateMaxKey
    Set Clone = cNodes
End Function

