VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRDiagnosises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection
Private mlngKey As Long

Public Property Get Item(vntIndexKey As Variant) As cEPRDiagnosis
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

Private Sub Class_initialize()
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
Public Function AddExistNode(vData As cEPRDiagnosis, Optional blnWithKeys As Boolean = False) As Long
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
    Dim objNewMember As cEPRDiagnosis
    Set objNewMember = New cEPRDiagnosis
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
'## 功能：  复制对象集合
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cEPRDiagnosises
    '复制对象
    Dim cNodes As New cEPRDiagnosises, lngKey As Long
    Dim i As Long
    For i = 1 To mCol.Count
        lngKey = cNodes.Add()
        If blnWithKeys Then
            cNodes("K" & lngKey).Key = mCol(i).Key
        End If
        cNodes("K" & lngKey).ID = mCol(i).ID
        cNodes("K" & lngKey).文件ID = mCol(i).文件ID
        cNodes("K" & lngKey).父ID = mCol(i).父ID
        cNodes("K" & lngKey).开始版 = mCol(i).开始版
        cNodes("K" & lngKey).终止版 = mCol(i).终止版
        cNodes("K" & lngKey).对象序号 = mCol(i).对象序号
        
        cNodes("K" & lngKey).类型 = mCol(i).类型
        cNodes("K" & lngKey).中医 = mCol(i).中医
        cNodes("K" & lngKey).疾病id = mCol(i).疾病id
        cNodes("K" & lngKey).诊断id = mCol(i).诊断id
        cNodes("K" & lngKey).证候id = mCol(i).证候id
        cNodes("K" & lngKey).描述 = mCol(i).描述
        cNodes("K" & lngKey).疑诊 = mCol(i).疑诊
        cNodes("K" & lngKey).日期 = mCol(i).日期
    Next i
    If blnWithKeys Then Call UpdateMaxKey
    Set Clone = cNodes
End Function

'################################################################################################################
'## 功能：  获取最大关键字（即，最近一次签名）
'################################################################################################################
Public Function GetMaxKey() As Long
    Dim i As Long, j As Long
    For i = 1 To mCol.Count
        If mCol(i).Key > j Then j = mCol(i).Key
    Next
    GetMaxKey = j
End Function

