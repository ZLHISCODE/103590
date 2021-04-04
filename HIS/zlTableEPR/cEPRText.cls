VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'多要素编辑时的文本
Option Explicit
Private mvarKey As Long                     '集合中的关键字
Private mvarID As Double                      '
Private mvar父ID As Double
Private mvar区域 As String
Private mvar保留对象 As Boolean
Private mvar内容行次 As Long
Private mvar内容文本 As String
Private mvar开始版 As Long
Private mvar终止版 As Long
Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Double)
    mvarID = vData
End Property

Public Property Get ID() As Double
    ID = mvarID
End Property
Public Property Let 父ID(ByVal vData As Double)
    mvar父ID = vData
End Property

Public Property Get 父ID() As Double
    父ID = mvar父ID
End Property

Public Property Let 区域(ByVal vData As String)
    mvar区域 = vData
End Property

Public Property Get 区域() As String
    区域 = mvar区域
End Property
Public Property Let 保留对象(ByVal vData As Boolean)
    mvar保留对象 = vData
End Property

Public Property Get 保留对象() As Boolean
    保留对象 = mvar保留对象
End Property

Public Property Let 内容行次(ByVal vData As Long)
    mvar内容行次 = vData
End Property

Public Property Get 内容行次() As Long
    内容行次 = mvar内容行次
End Property
Public Property Let 内容文本(ByVal vData As String)
    mvar内容文本 = vData
End Property

Public Property Get 内容文本() As String
    内容文本 = mvar内容文本
End Property
Public Property Let 开始版(ByVal vData As Long)
    mvar开始版 = vData
End Property

Public Property Get 开始版() As Long
    开始版 = mvar开始版
End Property
Public Property Let 终止版(ByVal vData As Long)
    mvar终止版 = vData
End Property

Public Property Get 终止版() As Long
    终止版 = mvar终止版
End Property
Public Sub Clone(cNode As cTabText)
    With cNode
        .保留对象 = .保留对象
        .内容行次 = Me.内容行次
        .内容文本 = Me.内容文本
        .开始版 = Me.开始版
        .终止版 = Me.终止版
    End With
End Sub
Public Function GetSaveSQL(ByVal editType As TabET, ByVal CellTmp As cTabCell, arrSQL As Variant) As Boolean
'功能:保存固定文本,单文本内容,混合区域的主项
    On Error GoTo errHand
    With CellTmp
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        Select Case editType
        Case TabET_病历文件定义
            arrSQL(UBound(arrSQL)) = "Zl_病历文件结构_Update(" & mvarID & "," & .文件ID & "," & .ID & _
                                        "," & .对象序号 & "," & cprCTText & ",Null," & IIf(mvar保留对象, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|'," & mvar内容行次 & ",'" & Replace(mvar内容文本, "'", "'" & "||Chr(39)||" & "'") & "')"
        Case TabET_全文示范编辑
            arrSQL(UBound(arrSQL)) = "Zl_病历范文内容_Update(" & mvarID & "," & .文件ID & "," & .ID & _
                                        "," & .对象序号 & "," & cprCTText & ",Null," & IIf(mvar保留对象, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|'," & mvar内容行次 & ",'" & Replace(mvar内容文本, "'", "'" & "||Chr(39)||" & "'") & "')"
        Case TabET_单病历编辑, TabET_单病历审核
            arrSQL(UBound(arrSQL)) = "Zl_电子病历内容_Update(" & mvarID & "," & .文件ID & "," & mvar开始版 & "," & mvar终止版 & "," & .ID & _
                                        "," & .对象序号 & "," & cprCTText & ",Null," & IIf(mvar保留对象, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|'," & mvar内容行次 & ",'" & Replace(mvar内容文本, "'", "'" & "||Chr(39)||" & "'") & "')"
        End Select
    End With
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function InsertIntoEditor(ByRef edtThis As Object, Optional ByVal lngStartPos As Long = -1) As Boolean
'******************************************************************************************************************
' 功能：  插入诊治要素到Editor中
'
' 参数：  edtThis         :当前的编辑器控件
'         lngStartPos     :当前位置
'         blnFirstIns     :是否是第一次插入（默认第一次插入会在末尾加逗号）
'******************************************************************************************************************
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '保证不能插入关键字内部
    
    If edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then Exit Function

    Dim p As Long, lngKey As Long, blnForce As Boolean, lngLen As Long, lngLen2 As Long, strTmp As String
    p = lngStartPos
    blnForce = edtThis.ForceEdit

    With edtThis
        .Freeze
        .ForceEdit = True
        .Range(lngStartPos, lngStartPos).Text = Me.内容文本
        If Me.保留对象 Then
            .Range(lngStartPos, Len(Me.内容文本)).Font.Protected = True '保护要素及隐藏区域
            .Range(lngStartPos, Len(Me.内容文本)).Font.BackColor = &HFFC0C0
        End If
        lngStartPos = lngStartPos + Len(Me.内容文本)                               '位置后移
        .Range(lngStartPos, lngStartPos).Selected                                           '光标定位
        .ForceEdit = blnForce
        .UnFreeze
    End With
    InsertIntoEditor = True
End Function
'################################################################################################################
'## 功能：  检查所有对象的ID是否有效
'##
'## 参数：  eEditType :当前的编辑模式
'##
'## 说明：  如果没有ID，则到数据库中提取一个唯一ID号。父ID=0的情况是单元素由Cell取ID
'################################################################################################################
Public Sub CheckValidID(eEditType As TabET)
    Dim i As Long, strTable As String
    Select Case eEditType
    Case TabET_病历文件定义
        strTable = "病历文件结构"
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId(strTable)
    Case TabET_全文示范编辑
        strTable = "病历范文内容"
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId(strTable)
    Case TabET_单病历编辑, TabET_单病历审核
        strTable = "电子病历内容"
        If Me.ID = 0 Then Me.ID = GetEPRContentNextId
    End Select
    
End Sub
