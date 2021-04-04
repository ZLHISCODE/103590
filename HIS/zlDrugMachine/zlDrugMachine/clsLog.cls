VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
'说明：日志类
'作者：YuZhiYong
'邮箱：yzyong@zlsoft.cn
'---------------------------------------------------------------

Private mobjFSO As FileSystemObject
Private mstrContent As String

Private mstrPath As String
Private mstrFileName As String
Private mintSaveDays As Integer
Private mblnAction As Boolean
Private mbytType As Byte

'日志类别；0-简要日志；1-详细日志
Public Property Get Type_() As Byte
    Type_ = mbytType
End Property
Public Property Let Type_(ByVal bytValue As Byte)
    mbytType = bytValue
End Property

'开启日志输出
Public Property Get Action() As Boolean
    Action = mblnAction
End Property
Public Property Let Action(ByVal blnValue As Boolean)
    mblnAction = blnValue
End Property

'存储效期
Public Property Get SaveDays() As Integer
    SaveDays = mintSaveDays
End Property
Public Property Let SaveDays(ByVal intValue As Integer)
    mintSaveDays = intValue
End Property

'日志文件名，无后缀名
Public Property Get FileName() As String
    FileName = mstrFileName
End Property
Public Property Let FileName(ByVal strValue As String)
    mstrFileName = strValue
End Property

'存储日志路径
Public Property Get Path() As String
    Path = mstrPath
End Property
Public Property Let Path(ByVal strValue As String)
    mstrPath = strValue
End Property

Public Sub Add(ByVal strContent As String, Optional ByVal bytLevel As Byte = 0, Optional ByVal bytType As Byte = 0)
'功能：暂时将日志内容保存至变量
'参数：
'  strContent：日志内容
'  bytLevel：层级文本叠加
'  bytType：日志类别；0-简要日志；1-详细日志

    If mblnAction = False Then Exit Sub

    If mbytType = 0 Then
        If bytType = 0 Then GoSub marAdd
    Else
        GoSub marAdd
    End If

    Exit Sub

marAdd:
    strContent = Replace(strContent, vbNewLine, vbNewLine & String(bytLevel * 2, vbKeySpace))
    mstrContent = mstrContent & String(bytLevel * 2, vbKeySpace) & strContent & vbNewLine
    Return
End Sub

Public Sub AddTime(Optional ByVal bytLevel As Byte = 0, Optional ByVal bytType As Byte = 0)
    Call Me.Add(vbNewLine & "" & Now, bytLevel, bytType)
End Sub

Public Sub Save()
'功能：正式将日志变量的内容写入日志文件
    
    Dim objTS As TextStream
    Dim objFolder As Folder
    Dim objFile As File
    Dim strDate As String, strFileName As String

    If mblnAction = False Then Exit Sub
    If Trim(mstrContent) = "" Then Exit Sub
    
    On Error GoTo hErr

    '自动生成日志文件
    
    strFileName = mstrFileName & Format(Date, "_yyyymmdd") & ".log"

    Set objTS = mobjFSO.OpenTextFile(mstrPath & "\" & strFileName, ForAppending, True)
    objTS.WriteLine Left(mstrContent, Len(mstrContent) - 1)   '日记内容
    objTS.Close
    
    ''检查x天外的日志文件，并删除
    Set objFolder = mobjFSO.GetFolder(mstrPath)
    For Each objFile In objFolder.Files
        If LCase(objFile.Name) Like LCase(mstrFileName) & "_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].log" Then
            strDate = Split(objFile.Name, "_")(1)
            strDate = Split(strDate, ".")(0)
            strDate = Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2)
            If Abs(Date - CDate(strDate)) >= mintSaveDays Then
                On Error Resume Next
                objFile.Delete True
                On Error GoTo hErr
            End If
        End If
    Next
    
    mstrContent = ""
    
    Exit Sub
    
hErr:
    mstrContent = ""
End Sub

Private Sub Class_Initialize()
    Set mobjFSO = New FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set mobjFSO = Nothing
End Sub
