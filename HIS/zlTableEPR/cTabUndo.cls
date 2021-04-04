VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarKey As String   '由"K" & Row_Col组成，表示所属的行列
Private mvarType As Byte    '单元格类型
Private mvarEkey As String  '要素Key串，可能是单要素，也可能是混合区域的要素串
Private mvarTxt As String   '固定文本或普通文本区域
Private mvarTkey As String  '文本串 混合区的文本串
Private mvarPkey As String  '图片关键字
Private mvarPmkey As String '图片标记关键字
Private mvarRow As Long     '所在行
Private mvarCol As Long     '所在列
Private mvarOrigPic As StdPicture           '原始图片
Public Property Let Key(ByVal vData As String)
    mvarKey = vData
    Me.Row = Split(Mid(mvarKey, 2), "_")(0)
    Me.Col = Split(Mid(mvarKey, 2), "_")(1)
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let CT(ByVal vData As Byte)
    mvarType = vData
End Property

Public Property Get CT() As Byte
    CT = mvarType
End Property
Public Property Let Ekey(ByVal vData As String)
    mvarEkey = vData
End Property

Public Property Get Ekey() As String
    Ekey = mvarEkey
End Property
Public Property Let CTxt(ByVal vData As String)
    mvarTxt = vData
End Property

Public Property Get CTxt() As String
    CTxt = mvarTxt
End Property
Public Property Let Tkey(ByVal vData As String)
    mvarTkey = vData
End Property

Public Property Get Tkey() As String
    Tkey = mvarTkey
End Property
Public Property Let PKey(ByVal vData As String)
    mvarPkey = vData
End Property

Public Property Get PKey() As String
    PKey = mvarPkey
End Property
Public Property Let PmKey(ByVal vData As String)
    mvarPmkey = vData
End Property

Public Property Get PmKey() As String
    PmKey = mvarPmkey
End Property
Public Property Let Row(ByVal vData As Long)
    mvarRow = vData
End Property
Public Property Get Row() As Long
    Row = mvarRow
End Property
Public Property Let Col(ByVal vData As Long)
    mvarCol = vData
End Property
Public Property Get Col() As Long
    Col = mvarCol
End Property
Public Property Set OrigPic(ByRef vData As StdPicture)
    Set mvarOrigPic = vData
End Property

Public Property Let OrigPic(ByRef vData As StdPicture)
    Set mvarOrigPic = vData
End Property

Public Property Get OrigPic() As StdPicture
    Set OrigPic = mvarOrigPic
End Property

Private Sub Class_Initialize()
    Set mvarOrigPic = New StdPicture
End Sub

Private Sub Class_Terminate()
    Set mvarOrigPic = Nothing
End Sub
