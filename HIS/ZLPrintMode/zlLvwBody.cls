VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlLvwBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'局部变量保存属性值
Private mvarobjData As Object  '本地复本
Private mvarFont As StdFont '本地复本

Public Property Set Font(ByVal vData As StdFont)
    Set mvarFont = vData
End Property


Public Property Get Font() As StdFont
    Set Font = mvarFont
End Property


Public Property Set objData(ByVal vData As Object)
'当把对象赋值给属性时在 Set 语句左边使用。
'Syntax: Set x.objData = Form1
    Set mvarobjData = vData
    mvarFont.Name = vData.Font.Name
    mvarFont.Size = vData.Font.Size
    mvarFont.Bold = vData.Font.Bold
    mvarFont.Italic = vData.Font.Italic
End Property

Public Property Get objData() As Object
Attribute objData.VB_UserMemId = 0
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.objData
    Set objData = mvarobjData
End Property


Private Sub Class_Initialize()
    Set mvarFont = New StdFont
    mvarFont.Name = "宋体"
    mvarFont.Size = 9
    mvarFont.Bold = False
    mvarFont.Italic = False
End Sub

Private Sub Class_Terminate()
    Set mvarFont = Nothing
    Set mvarobjData = Nothing
End Sub
