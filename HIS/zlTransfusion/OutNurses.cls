VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutNurses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"OutNurse"
Attribute VB_Ext_KEY = "Member0" ,"OutNurse"
Option Explicit

'局部变量，保存集合
Private mCol As Collection


Public Function getOutNurse(ByVal lngDeptID As Long) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strName As String
    On Error GoTo hErr
    
    strSQL = "Select /*+Rule */ Distinct D.姓名 " & vbNewLine & _
            "From 人员表 D, 人员性质说明 B, 部门人员 A" & vbNewLine & _
            "Where A.人员id = D.ID And A.人员id = B.人员id And (D.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or D.撤档时间 Is Null) And B.人员性质 = '护士' And A.部门id = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select Distinct D.姓名 " & vbNewLine & _
            "From 人员表 D, 人员性质说明 C, 部门人员 B," & vbNewLine & _
            "     (Select A.ID From 部门表 A Start With 上级id = [1] Connect By Prior ID = 上级id) A" & vbNewLine & _
            "Where B.人员id = D.ID And A.ID = B.部门id And B.人员id = C.人员id And (D.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or D.撤档时间 Is Null) And C.人员性质 = '护士'"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "outNurses.getOutNurse", lngDeptID)
    Call ClearOutNurse
    Do Until rsTmp.EOF
        strName = zlCommFun.NVL(rsTmp.Fields("姓名"))
        Add strName, strName
        rsTmp.MoveNext
    Loop
    Exit Function
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Add(姓名 As String, Optional sKey As String) As OutNurse
    '创建新对象
    Dim objNewMember As OutNurse
    Set objNewMember = New OutNurse


    '设置传入方法的属性
    objNewMember.姓名 = 姓名
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As OutNurse
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    ' x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Public Sub ClearOutNurse()
    Dim lngIndex As Long
    
    For lngIndex = 1 To Me.Count
        Remove 1
    Next
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub

