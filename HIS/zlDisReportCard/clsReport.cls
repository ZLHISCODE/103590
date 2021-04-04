VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrID As String
Private mstrFileID As String
Private mstrStartR As String
Private mstrStopR As String
Private mstrObjNo As String
Private mstrObjType As String
Private mstrTxt As String
Private mstrReplace As String
Private mstrEleName As String
Private mstrEleType As String
Private mstrEleIdt As String
Private mstrEleRange As String

Public Function GetSaveSql(arrSql() As Variant, blnAddSql As Boolean) As Boolean
    Dim intIndex As Integer
    
    On Error GoTo errHand
    If blnAddSql = True Then
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        intIndex = UBound(arrSql)
    Else
        intIndex = val(mstrObjNo) - 1
    End If
    If val(mstrID) = 0 Then
        mstrID = GetNextDoubleId("电子病历内容")
    End If
    arrSql(intIndex) = "Zl_传染病报告卡内容_Update(" & mstrID & "," & mstrFileID & "," & mstrStartR & "," & mstrStopR & _
                        "," & mstrObjNo & "," & mstrObjType & ",'" & mstrTxt & "'," & mstrReplace & ",'" & mstrEleName & _
                        "'," & mstrEleType & "," & mstrEleIdt & ",'" & mstrEleRange & "')"
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

'ID
Property Get ID() As String
    ID = mstrID
End Property

Property Let ID(ByVal vNewValue As String)
    mstrID = vNewValue
End Property

'文件ID
Property Get FileID() As String
    FileID = mstrFileID
End Property

Property Let FileID(ByVal vNewValue As String)
    mstrFileID = vNewValue
End Property

'开始版本
Property Get StartR() As String
    StartR = mstrStartR
End Property

Property Let StartR(ByVal vNewValue As String)
    mstrStartR = vNewValue
End Property

'终止版本
Property Get StopR() As String
    StopR = mstrStopR
End Property

Property Let StopR(ByVal vNewValue As String)
    mstrStopR = vNewValue
End Property

'对象序号
Property Get ObjNo() As String
    ObjNo = mstrObjNo
End Property

Property Let ObjNo(ByVal vNewValue As String)
    mstrObjNo = vNewValue
End Property

'对象类型
Property Get ObjType() As String
    ObjType = mstrObjType
End Property

Property Let ObjType(ByVal vNewValue As String)
    mstrObjType = vNewValue
End Property

'内容文本
Property Get Txt() As String
    Txt = mstrTxt
End Property

Property Let Txt(ByVal vNewValue As String)
    mstrTxt = vNewValue
End Property

'替换域
Property Get Replace() As String
    Replace = mstrReplace
End Property

Property Let Replace(ByVal vNewValue As String)
    mstrReplace = vNewValue
End Property

'要素名称
Property Get EleName() As String
    EleName = mstrEleName
End Property

Property Let EleName(ByVal vNewValue As String)
    mstrEleName = vNewValue
End Property

'要素类型
Property Get EleType() As String
    EleType = mstrEleType
End Property

Property Let EleType(ByVal vNewValue As String)
    mstrEleType = vNewValue
End Property

'要素表示
Property Get EleIdt() As String
    EleIdt = mstrEleIdt
End Property

Property Let EleIdt(ByVal vNewValue As String)
    mstrEleIdt = vNewValue
End Property

'要素范围
Property Get EleRange() As String
    EleRange = mstrEleRange
End Property

Property Let EleRange(ByVal vNewValue As String)
    mstrEleRange = vNewValue
End Property

