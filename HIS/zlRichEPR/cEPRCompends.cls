VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRCompends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cEPRCompend"
Attribute VB_Ext_KEY = "Member0" ,"cEPRCompend"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlngKey As Long     '流水号

Public Property Get Item(vntIndexKey As Variant) As cEPRCompend
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
'## 功能：  添加一个已经存在的提纲对象到集合中
'##
'## 参数：  blnWithKeys :表示是否同时复制Key值
'################################################################################################################
Public Function AddExistNode(vData As cEPRCompend, Optional blnWithKeys As Boolean = False) As Long
    Dim strKey As String
    If blnWithKeys = False Then
        mlngKey = mlngKey + 1
        strKey = "K" & mlngKey
    Else
        strKey = "K" & vData.Key
    End If
    mCol.Add vData.Clone(blnWithKeys), strKey
        
    UpdateParentKey mCol(strKey)
    
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
    Dim objNewMember As cEPRCompend
    Set objNewMember = New cEPRCompend
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
'## 功能：  根据文件ID将提纲保存到数据库中
'##
'## 参数：  eEditType :当前的编辑模式
'##         lngFileId :文件ID
'##
'## 说明：  必须先获取提纲的有效顺序号，才能保证各提纲关系的正确性
'################################################################################################################
Public Function SaveCompendsToDB(eEdtType As EditTypeEnum, lngFileID As Long) As Boolean
    Err = 0: On Error GoTo LL
    Dim i As Long, j As Long, k As Long
    
    Call CheckValidIDs(eEdtType)
    Call UpdateParentIDs

    For i = 1 To mCol.Count
        gstrSQL = mCol(i).SaveCompendToDB(eEdtType)             '获取该提纲保存的SQL串！～
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, "cEPRCompends")
    Next
    SaveCompendsToDB = True
    Exit Function
LL:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    SaveCompendsToDB = False
End Function

'################################################################################################################
'## 功能：  直接从数据库获取提纲到集合中
'##
'## 参数：  eEditType :当前的编辑模式
'##         lngFileId :文件ID
'################################################################################################################
Public Function GetCompendsFromDB(eEdtType As EditTypeEnum, lngFileID As Long) As Boolean
'    On Error GoTo LL
'    Dim lngKey As Long
'
'    Select Case eEdtType
'    Case cprET_病历文件定义
'        gstrSQL = "Select Level, t.* " & _
'            "From 病历文件结构 t " & _
'            "Where 文件ID = [1] And 对象类型 = 1 " & _
'            "Start With 父id Is Null " & _
'            "Connect By Prior Id = 父id " & _
'            "Order By Level, 对象序号"
'    Case cprET_全文示范编辑
'        gstrSQL = "Select Level, t.* " & _
'            "From 病历范文内容 t " & _
'            "Where 文件ID = [1] And 对象类型 = 1 " & _
'            "Start With 父id Is Null " & _
'            "Connect By Prior Id = 父id " & _
'            "Order By Level, 对象序号"
'    Case cprET_单病历编辑, cprET_单病历审核
'        gstrSQL = "Select Level, t.* " & _
'            "From 电子病历内容 t " & _
'            "Where 文件ID = [1] And 对象类型 = 1 " & _
'            "Start With 父id Is Null " & _
'            "Connect By Prior Id = 父id " & _
'            "Order By Level, 对象序号"
'    End Select
'    Dim Rs As New ADODB.Recordset
'    Rs.CursorLocation = adUseClient
'
'    Call SQLTest(App.ProductName, "cEPRCompends", gstrSQL)
'    Set Rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRCompends", lngFileID)
'    Call SQLTest
'
'    If Not Rs.EOF Then
'        '读取提纲
'        Set mCol = New Collection
'        mlngKey = 0
'        Do While Not Rs.EOF
'            lngKey = Add(NVL(Rs("对象标记"), 0))       '恢复Key值！
'
'            Me("K" & lngKey).ID = Rs("ID")
'            Me("K" & lngKey).文件ID = NVL(Rs("文件ID"), 0)
'            Me("K" & lngKey).父ID = NVL(Rs("父ID"), 0)
'            Me("K" & lngKey).对象序号 = NVL(Rs("对象序号"), 0)
'            Me("K" & lngKey).保留对象 = IIf(NVL(Rs("保留对象"), 0) = 0, False, True)
'            Me("K" & lngKey).名称 = NVL(Rs("内容文本"))
'            Me("K" & lngKey).说明 = NVL(Rs("对象属性"))
'            Me("K" & lngKey).预制提纲ID = NVL(Rs("预制提纲ID"), 0)
'            Me("K" & lngKey).复用提纲 = IIf(NVL(Rs("复用提纲"), 0) = 0, False, True)
'            Me("K" & lngKey).使用时机 = NVL(Rs("使用时机"))
'            Me("K" & lngKey).Level = NVL(Rs("Level"), 1)
'            Me("K" & lngKey).内部序号 = Me("K" & lngKey).对象序号
'
'            If eEdtType = cprET_单病历审核 Or eEdtType = cprET_单病历编辑 Then
'                Me("K" & lngKey).版本号 = NVL(Rs("版本号"), 0)
'            End If
'
'            Rs.MoveNext
'        Loop
'
'        Call UpdateParentKeys               '更新父Key
'
'        GetCompendsFromDB = True
'    Else
'        GetCompendsFromDB = False
'    End If
'    Exit Function
'LL:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'    GetCompendsFromDB = False
End Function

'################################################################################################################
'## 功能：  根据父ID，更新指定节点的父Key
'##
'## 参数：  vData : 需要更新的节点
'################################################################################################################
Public Sub UpdateParentKey(vData As cEPRCompend)
    '根据父ID，更新父Key，从而确定上下级关系。
    Dim i As Long
    For i = 1 To mCol.Count
        If mCol(i).ID = vData.父ID And vData.父ID <> 0 Then vData.父Key = i: Exit For
    Next
End Sub

'################################################################################################################
'## 功能：  根据父ID，更新所有节点的父Key
'################################################################################################################
Public Sub UpdateParentKeys()
    Dim i As Long, j As Long
    For i = 1 To mCol.Count
        For j = 1 To mCol.Count
            If mCol(i).父ID = mCol(j).ID And mCol(i).父ID <> 0 Then
                mCol(i).父Key = mCol(j).Key
                mCol(i).Level = mCol(j).Level + 1
                Exit For
            End If
        Next
    Next
End Sub

'################################################################################################################
'## 功能：  检查父Key是否有效（比如在删除某提纲后，它的下级提纲应该父Key＝0）
'################################################################################################################
Public Sub CheckValidParentKeys()
    Dim i As Long, j As Long, lng父Key As Long
    For i = 1 To mCol.Count
        lng父Key = mCol(i).父Key
        mCol(i).父Key = 0
        For j = 1 To mCol.Count
            If lng父Key = mCol(j).Key And lng父Key <> 0 Then
                mCol(i).父Key = mCol(j).Key
                mCol(i).Level = mCol(j).Level + 1
                Exit For
            End If
        Next
        If mCol(i).父Key = 0 Then mCol(i).Level = 1 '删除根提纲后,子提纲自动并入一级提纲
    Next
End Sub

'################################################################################################################
'## 功能：  根据父Key，更新所有节点的父ID
'################################################################################################################
Public Sub UpdateParentIDs()
    Dim i As Long, j As Long
    For i = 1 To mCol.Count
        For j = 1 To mCol.Count
            If mCol(i).父Key = mCol(j).Key Then
                mCol(i).父ID = mCol(j).ID
                Exit For
            End If
        Next
        If mCol(i).父Key = 0 Then mCol(i).父ID = 0 '根提纲没有父ID
    Next
End Sub

'################################################################################################################
'## 功能：  将节点填充到提纲树中
'##
'## 参数：  objTree     :需要填充的TreeView控件
'##         lngSelKey   :默认选中的节点
'##
'## 说明：  必须先获取提纲的有效顺序号，才能保证填充的正确性
'################################################################################################################
Public Sub FillTree(ByVal objTree As Object, Optional lngSelKey As Long = 0)
    Dim i As Long, j As Long, k As Long
    Dim lngMaxOrder As Long
    
    objTree.Nodes.Clear
    If mCol.Count = 0 Then Exit Sub
    
    LockWindowUpdate objTree.hWnd
    Dim Node As Node
    '先找最大序号
    For i = 1 To mCol.Count
        If mCol(i).内部序号 > lngMaxOrder Then lngMaxOrder = mCol(i).内部序号
    Next
    For i = 1 To lngMaxOrder
        '按照序号由小到大的循序来加入根节点。
        For j = 1 To mCol.Count
            If mCol(j).父Key = 0 And mCol(j).内部序号 = i Then
                Set Node = objTree.Nodes.Add(Null, tvwChild, "K" & mCol(j).Key, mCol(j).名称, 1, 1)
                If lngSelKey = mCol(j).Key Then Node.Selected = True
                Node.Tag = mCol(j).Key
                Node.Expanded = True
'                Debug.Print "内部序号:" & mCol(j).内部序号 & vbTab & "Key:" & Node.Key & vbTab & "Tag:" & _
'                    Node.Tag & vbTab & "名称:" & mCol(j).名称 & vbTab & "父Key:" & mCol(j).父Key
                AddChildNodes objTree, Node, mCol(j).Key  '递归加入下级子节点。
                Exit For
            End If
        Next
    Next

    LockWindowUpdate 0
    UpdateWindow objTree.hWnd
End Sub

'################################################################################################################
'## 功能：  递归添加本节点的所有子节点
'##
'## 参数：  objTree :需要填充的TreeView控件
'##         Node    :上级的树节点
'##         lngKey  :上级的关键字
'##
'## 说明：  必须先获取提纲的有效顺序号，才能保证填充的正确性
'################################################################################################################
Private Sub AddChildNodes(ByVal objTree As TreeView, ByVal Node As Node, ByVal lngKey As Long)
    Dim i As Long, SubNode As Node, j As Long, k As Long
    Dim lngMaxOrder As Long
    '先找最大序号
    For i = 1 To mCol.Count
        If mCol(i).内部序号 > lngMaxOrder Then lngMaxOrder = mCol(i).内部序号
    Next
    For i = 1 To lngMaxOrder
        '按照序号由小到大的循序来加入根节点。
        For j = 1 To mCol.Count
            If mCol(j).父Key = lngKey And mCol(j).内部序号 = i Then
                Set SubNode = objTree.Nodes.Add(Node, tvwChild, "K" & mCol(j).Key, mCol(j).名称, 2, 2)
                SubNode.Tag = mCol(j).Key
                SubNode.Expanded = True
                AddChildNodes objTree, SubNode, mCol(j).Key
'                Debug.Print "内部序号:" & mCol(j).内部序号 & vbTab & "Key:" & SubNode.Key & vbTab & "Tag:" & _
'                    SubNode.Tag & vbTab & "名称:" & mCol(j).名称 & vbTab & "父Key:" & mCol(j).父Key
                Exit For
            End If
        Next
    Next
End Sub

'################################################################################################################
'## 功能：  递归找出指定提纲的根提纲的Key值
'##
'## 参数：  lngKey  :需要查找的节点关键字
'################################################################################################################
Public Function GetRootNodeKey(ByVal lngKey As Long) As Long
    Dim i As Long, lngR As Long
    If mCol("K" & lngKey).父Key = 0 Then
        lngR = 0
    Else
        For i = 1 To mCol.Count
            If mCol(i).Key = mCol("K" & lngKey).父Key Then
                lngR = i
                If mCol(i).Level > 1 Then
                    lngR = GetParentNodeKey(mCol(i).Key)
                Else
                    Exit For
                End If
            End If
        Next
    End If
    GetRootNodeKey = lngR
End Function

'################################################################################################################
'## 功能：  找出指定提纲的上级提纲
'##
'## 参数：  lngKey  :需要查找的节点关键字
'################################################################################################################
Public Function GetParentNodeKey(ByVal lngKey As Long) As Long
    Dim i As Long, lngR As Long
    If mCol("K" & lngKey).父Key = 0 Then
        lngR = 0
    Else
        For i = 1 To mCol.Count
            If mCol(i).Key = mCol("K" & lngKey).父Key Then
                lngR = mCol(i).Key
                Exit For
            End If
        Next
    End If
    GetParentNodeKey = lngR
End Function

'################################################################################################################
'## 功能：  从文本中更新提纲序号
'##
'## 参数：  edtThis  :RTF编辑器控件
'##
'## 说明：  可能影响到树节点的排列
'################################################################################################################
Public Function UpdateOrdersFromText(edtThis As Object) As Boolean
    On Error GoTo LL
    Dim i As Long, j As Long, lKey As Long
    Dim lngOrder As Long
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！

    lngOrder = 0
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = 1
LL1:
        Do While i > 0
            i = InStr(i, sText, "OS(", vbTextCompare)   '反向搜索第一个 OS(...) 关键字
            If i <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                    i = i + 1
                    GoTo LL1
                End If
                '已找到起始关键字
                '查找结束关键字
                j = i + 16
LL2:
                j = InStr(j, sText, "OE(", vbTextCompare)   '正向搜索其匹配的结束关键字 OE(...)
                If j <> 0 Then
                    '看是否是关键字
                    If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                        j = j + 1
                        GoTo LL2
                    End If
                    '找到结束关键字
                    lSS = i - 1
                    lSE = lSS + 16
                    lES = j - 1
                    lEE = lES + 16
                    lKey = Val(.TOM.TextDocument.Range(lSS + 3, lSS + 11))
                    lngOrder = lngOrder + 1

                    mCol("K" & lKey).内部序号 = lngOrder
                    i = lEE
                End If
            End If
        Loop
    End With
    
    UpdateOrdersFromText = True
    Exit Function
LL:
    UpdateOrdersFromText = False
End Function

