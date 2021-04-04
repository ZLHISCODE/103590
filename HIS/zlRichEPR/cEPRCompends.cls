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
Private mlngKey As Long     '��ˮ��

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
'## ���ܣ�  ����һ���Ѿ����ڵ���ٶ��󵽼�����
'##
'## ������  blnWithKeys :��ʾ�Ƿ�ͬʱ����Keyֵ
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
'## ���ܣ�  �����¶���
'##
'## ������  lngKey :���ڻָ��ɵļ��϶��󣬲��������Ҫ���������ˮ�š�
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
'## ���ܣ�  ������ж����ID�Ƿ���Ч
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š�
'################################################################################################################
Public Sub CheckValidIDs(eEditType As EditTypeEnum)
    Dim i As Long
    Select Case eEditType
    Case cprET_�����ļ�����
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = zlDatabase.GetNextId("�����ļ��ṹ")
        Next
    Case cprET_ȫ��ʾ���༭
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = zlDatabase.GetNextId("������������")
        Next
    Case cprET_�������༭, cprET_���������
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = GetEPRContentNextId
        Next
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  ���������ˮ��
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
'## ���ܣ�  �����ļ�ID����ٱ��浽���ݿ���
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##         lngFileId :�ļ�ID
'##
'## ˵����  �����Ȼ�ȡ��ٵ���Ч˳��ţ����ܱ�֤����ٹ�ϵ����ȷ��
'################################################################################################################
Public Function SaveCompendsToDB(eEdtType As EditTypeEnum, lngFileID As Long) As Boolean
    Err = 0: On Error GoTo LL
    Dim i As Long, j As Long, k As Long
    
    Call CheckValidIDs(eEdtType)
    Call UpdateParentIDs

    For i = 1 To mCol.Count
        gstrSQL = mCol(i).SaveCompendToDB(eEdtType)             '��ȡ����ٱ����SQL������
        
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
'## ���ܣ�  ֱ�Ӵ����ݿ��ȡ��ٵ�������
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##         lngFileId :�ļ�ID
'################################################################################################################
Public Function GetCompendsFromDB(eEdtType As EditTypeEnum, lngFileID As Long) As Boolean
'    On Error GoTo LL
'    Dim lngKey As Long
'
'    Select Case eEdtType
'    Case cprET_�����ļ�����
'        gstrSQL = "Select Level, t.* " & _
'            "From �����ļ��ṹ t " & _
'            "Where �ļ�ID = [1] And �������� = 1 " & _
'            "Start With ��id Is Null " & _
'            "Connect By Prior Id = ��id " & _
'            "Order By Level, �������"
'    Case cprET_ȫ��ʾ���༭
'        gstrSQL = "Select Level, t.* " & _
'            "From ������������ t " & _
'            "Where �ļ�ID = [1] And �������� = 1 " & _
'            "Start With ��id Is Null " & _
'            "Connect By Prior Id = ��id " & _
'            "Order By Level, �������"
'    Case cprET_�������༭, cprET_���������
'        gstrSQL = "Select Level, t.* " & _
'            "From ���Ӳ������� t " & _
'            "Where �ļ�ID = [1] And �������� = 1 " & _
'            "Start With ��id Is Null " & _
'            "Connect By Prior Id = ��id " & _
'            "Order By Level, �������"
'    End Select
'    Dim Rs As New ADODB.Recordset
'    Rs.CursorLocation = adUseClient
'
'    Call SQLTest(App.ProductName, "cEPRCompends", gstrSQL)
'    Set Rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRCompends", lngFileID)
'    Call SQLTest
'
'    If Not Rs.EOF Then
'        '��ȡ���
'        Set mCol = New Collection
'        mlngKey = 0
'        Do While Not Rs.EOF
'            lngKey = Add(NVL(Rs("������"), 0))       '�ָ�Keyֵ��
'
'            Me("K" & lngKey).ID = Rs("ID")
'            Me("K" & lngKey).�ļ�ID = NVL(Rs("�ļ�ID"), 0)
'            Me("K" & lngKey).��ID = NVL(Rs("��ID"), 0)
'            Me("K" & lngKey).������� = NVL(Rs("�������"), 0)
'            Me("K" & lngKey).�������� = IIf(NVL(Rs("��������"), 0) = 0, False, True)
'            Me("K" & lngKey).���� = NVL(Rs("�����ı�"))
'            Me("K" & lngKey).˵�� = NVL(Rs("��������"))
'            Me("K" & lngKey).Ԥ�����ID = NVL(Rs("Ԥ�����ID"), 0)
'            Me("K" & lngKey).������� = IIf(NVL(Rs("�������"), 0) = 0, False, True)
'            Me("K" & lngKey).ʹ��ʱ�� = NVL(Rs("ʹ��ʱ��"))
'            Me("K" & lngKey).Level = NVL(Rs("Level"), 1)
'            Me("K" & lngKey).�ڲ���� = Me("K" & lngKey).�������
'
'            If eEdtType = cprET_��������� Or eEdtType = cprET_�������༭ Then
'                Me("K" & lngKey).�汾�� = NVL(Rs("�汾��"), 0)
'            End If
'
'            Rs.MoveNext
'        Loop
'
'        Call UpdateParentKeys               '���¸�Key
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
'## ���ܣ�  ���ݸ�ID������ָ���ڵ�ĸ�Key
'##
'## ������  vData : ��Ҫ���µĽڵ�
'################################################################################################################
Public Sub UpdateParentKey(vData As cEPRCompend)
    '���ݸ�ID�����¸�Key���Ӷ�ȷ�����¼���ϵ��
    Dim i As Long
    For i = 1 To mCol.Count
        If mCol(i).ID = vData.��ID And vData.��ID <> 0 Then vData.��Key = i: Exit For
    Next
End Sub

'################################################################################################################
'## ���ܣ�  ���ݸ�ID���������нڵ�ĸ�Key
'################################################################################################################
Public Sub UpdateParentKeys()
    Dim i As Long, j As Long
    For i = 1 To mCol.Count
        For j = 1 To mCol.Count
            If mCol(i).��ID = mCol(j).ID And mCol(i).��ID <> 0 Then
                mCol(i).��Key = mCol(j).Key
                mCol(i).Level = mCol(j).Level + 1
                Exit For
            End If
        Next
    Next
End Sub

'################################################################################################################
'## ���ܣ�  ��鸸Key�Ƿ���Ч��������ɾ��ĳ��ٺ������¼����Ӧ�ø�Key��0��
'################################################################################################################
Public Sub CheckValidParentKeys()
    Dim i As Long, j As Long, lng��Key As Long
    For i = 1 To mCol.Count
        lng��Key = mCol(i).��Key
        mCol(i).��Key = 0
        For j = 1 To mCol.Count
            If lng��Key = mCol(j).Key And lng��Key <> 0 Then
                mCol(i).��Key = mCol(j).Key
                mCol(i).Level = mCol(j).Level + 1
                Exit For
            End If
        Next
        If mCol(i).��Key = 0 Then mCol(i).Level = 1 'ɾ������ٺ�,������Զ�����һ�����
    Next
End Sub

'################################################################################################################
'## ���ܣ�  ���ݸ�Key���������нڵ�ĸ�ID
'################################################################################################################
Public Sub UpdateParentIDs()
    Dim i As Long, j As Long
    For i = 1 To mCol.Count
        For j = 1 To mCol.Count
            If mCol(i).��Key = mCol(j).Key Then
                mCol(i).��ID = mCol(j).ID
                Exit For
            End If
        Next
        If mCol(i).��Key = 0 Then mCol(i).��ID = 0 '�����û�и�ID
    Next
End Sub

'################################################################################################################
'## ���ܣ�  ���ڵ���䵽�������
'##
'## ������  objTree     :��Ҫ����TreeView�ؼ�
'##         lngSelKey   :Ĭ��ѡ�еĽڵ�
'##
'## ˵����  �����Ȼ�ȡ��ٵ���Ч˳��ţ����ܱ�֤������ȷ��
'################################################################################################################
Public Sub FillTree(ByVal objTree As Object, Optional lngSelKey As Long = 0)
    Dim i As Long, j As Long, k As Long
    Dim lngMaxOrder As Long
    
    objTree.Nodes.Clear
    If mCol.Count = 0 Then Exit Sub
    
    LockWindowUpdate objTree.hWnd
    Dim Node As Node
    '����������
    For i = 1 To mCol.Count
        If mCol(i).�ڲ���� > lngMaxOrder Then lngMaxOrder = mCol(i).�ڲ����
    Next
    For i = 1 To lngMaxOrder
        '���������С�����ѭ����������ڵ㡣
        For j = 1 To mCol.Count
            If mCol(j).��Key = 0 And mCol(j).�ڲ���� = i Then
                Set Node = objTree.Nodes.Add(Null, tvwChild, "K" & mCol(j).Key, mCol(j).����, 1, 1)
                If lngSelKey = mCol(j).Key Then Node.Selected = True
                Node.Tag = mCol(j).Key
                Node.Expanded = True
'                Debug.Print "�ڲ����:" & mCol(j).�ڲ���� & vbTab & "Key:" & Node.Key & vbTab & "Tag:" & _
'                    Node.Tag & vbTab & "����:" & mCol(j).���� & vbTab & "��Key:" & mCol(j).��Key
                AddChildNodes objTree, Node, mCol(j).Key  '�ݹ�����¼��ӽڵ㡣
                Exit For
            End If
        Next
    Next

    LockWindowUpdate 0
    UpdateWindow objTree.hWnd
End Sub

'################################################################################################################
'## ���ܣ�  �ݹ����ӱ��ڵ�������ӽڵ�
'##
'## ������  objTree :��Ҫ����TreeView�ؼ�
'##         Node    :�ϼ������ڵ�
'##         lngKey  :�ϼ��Ĺؼ���
'##
'## ˵����  �����Ȼ�ȡ��ٵ���Ч˳��ţ����ܱ�֤������ȷ��
'################################################################################################################
Private Sub AddChildNodes(ByVal objTree As TreeView, ByVal Node As Node, ByVal lngKey As Long)
    Dim i As Long, SubNode As Node, j As Long, k As Long
    Dim lngMaxOrder As Long
    '����������
    For i = 1 To mCol.Count
        If mCol(i).�ڲ���� > lngMaxOrder Then lngMaxOrder = mCol(i).�ڲ����
    Next
    For i = 1 To lngMaxOrder
        '���������С�����ѭ����������ڵ㡣
        For j = 1 To mCol.Count
            If mCol(j).��Key = lngKey And mCol(j).�ڲ���� = i Then
                Set SubNode = objTree.Nodes.Add(Node, tvwChild, "K" & mCol(j).Key, mCol(j).����, 2, 2)
                SubNode.Tag = mCol(j).Key
                SubNode.Expanded = True
                AddChildNodes objTree, SubNode, mCol(j).Key
'                Debug.Print "�ڲ����:" & mCol(j).�ڲ���� & vbTab & "Key:" & SubNode.Key & vbTab & "Tag:" & _
'                    SubNode.Tag & vbTab & "����:" & mCol(j).���� & vbTab & "��Key:" & mCol(j).��Key
                Exit For
            End If
        Next
    Next
End Sub

'################################################################################################################
'## ���ܣ�  �ݹ��ҳ�ָ����ٵĸ���ٵ�Keyֵ
'##
'## ������  lngKey  :��Ҫ���ҵĽڵ�ؼ���
'################################################################################################################
Public Function GetRootNodeKey(ByVal lngKey As Long) As Long
    Dim i As Long, lngR As Long
    If mCol("K" & lngKey).��Key = 0 Then
        lngR = 0
    Else
        For i = 1 To mCol.Count
            If mCol(i).Key = mCol("K" & lngKey).��Key Then
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
'## ���ܣ�  �ҳ�ָ����ٵ��ϼ����
'##
'## ������  lngKey  :��Ҫ���ҵĽڵ�ؼ���
'################################################################################################################
Public Function GetParentNodeKey(ByVal lngKey As Long) As Long
    Dim i As Long, lngR As Long
    If mCol("K" & lngKey).��Key = 0 Then
        lngR = 0
    Else
        For i = 1 To mCol.Count
            If mCol(i).Key = mCol("K" & lngKey).��Key Then
                lngR = mCol(i).Key
                Exit For
            End If
        Next
    End If
    GetParentNodeKey = lngR
End Function

'################################################################################################################
'## ���ܣ�  ���ı��и���������
'##
'## ������  edtThis  :RTF�༭���ؼ�
'##
'## ˵����  ����Ӱ�쵽���ڵ������
'################################################################################################################
Public Function UpdateOrdersFromText(edtThis As Object) As Boolean
    On Error GoTo LL
    Dim i As Long, j As Long, lKey As Long
    Dim lngOrder As Long
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��

    lngOrder = 0
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = 1
LL1:
        Do While i > 0
            i = InStr(i, sText, "OS(", vbTextCompare)   '����������һ�� OS(...) �ؼ���
            If i <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                    i = i + 1
                    GoTo LL1
                End If
                '���ҵ���ʼ�ؼ���
                '���ҽ����ؼ���
                j = i + 16
LL2:
                j = InStr(j, sText, "OE(", vbTextCompare)   '����������ƥ��Ľ����ؼ��� OE(...)
                If j <> 0 Then
                    '���Ƿ��ǹؼ���
                    If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                        j = j + 1
                        GoTo LL2
                    End If
                    '�ҵ������ؼ���
                    lSS = i - 1
                    lSE = lSS + 16
                    lES = j - 1
                    lEE = lES + 16
                    lKey = Val(.TOM.TextDocument.Range(lSS + 3, lSS + 11))
                    lngOrder = lngOrder + 1

                    mCol("K" & lKey).�ڲ���� = lngOrder
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
