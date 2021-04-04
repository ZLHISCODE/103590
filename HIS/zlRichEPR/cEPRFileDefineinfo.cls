VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRFileDefineInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'病历文件定义
Public ID As Long               'ID
Public 种类 As EPRDocTypeEnum   '种类：1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-诊断文书;6-知情文件;7-诊疗报告;8-诊疗申请
Public 编号 As Long             '文件编号
Public 名称 As String           '文件名称
Public 说明 As String           '文件说明
Public 页面 As String           '文件页面：只有住院病历和护理记录可以设置页面,其他类型的病历独立成页,页面编号和文件编号一致
Public 保留 As Long             '是否保留：0-用户定义的病历文件;1-系统保留的普通病历文件,不能删除;<0,指单独编制程序实现的特殊病历,不能删除
Public 通用 As Long             '通用性：0-未用;1-通用;2-专科
Public 页面名称 As String       '病历页面名称

Public 报表 As Long             '对诊疗单据:0-可编辑,并按编辑格式输出;1-无编辑,用自定义报表输出(如处方笺);2-可编辑,用自定义报表输出(如各类检查报告)；对护理记录:0-病人记录;1-病房记录。其他无效。
Public 格式 As String           'PaperKind;PaperOrient;PaperHeight;PaperWidth;MarginLeft;MarginRight;MarginTop;MarginBottom;BackColor;PaperColor;ShowPageNumber
Public 页眉 As String
Public 页脚 As String

Public lngModule As Long

Private HeadFontName As String
Private HeadFontSize As Long
Private HeadFontBold As Boolean
Private HeadFontItalic As Boolean
Private HeadFontUnderline As Boolean
Private HeadFontStrikethrough As Boolean
Private HeadFontColor As OLE_COLOR

Private FootFontName As String
Private FootFontSize As Long
Private FootFontBold As Boolean
Private FootFontItalic As Boolean
Private FootFontUnderline As Boolean
Private FootFontStrikethrough As Boolean
Private FootFontColor As OLE_COLOR

Public Property Let HeadFont(vData As String)
    On Error Resume Next
    Dim T As Variant
    T = Split(vData, "|")
    HeadFontName = T(0)
    HeadFontSize = T(1)
    HeadFontBold = T(2)
    HeadFontItalic = T(3)
    HeadFontUnderline = T(4)
    HeadFontStrikethrough = T(5)
    HeadFontColor = T(6)
End Property

Public Property Get HeadFont() As String
    HeadFont = HeadFontName & "|" & HeadFontSize & "|" & HeadFontBold & "|" & HeadFontItalic & "|" & HeadFontUnderline & "|" & HeadFontStrikethrough & "|" & HeadFontColor
End Property

Public Property Let FootFont(vData As String)
    On Error Resume Next
    Dim T As Variant
    T = Split(vData, "|")
    FootFontName = T(0)
    FootFontSize = T(1)
    FootFontBold = T(2)
    FootFontItalic = T(3)
    FootFontUnderline = T(4)
    FootFontStrikethrough = T(5)
    FootFontColor = T(6)
End Property

Public Property Get FootFont() As String
    FootFont = FootFontName & "|" & FootFontSize & "|" & FootFontBold & "|" & FootFontItalic & "|" & FootFontUnderline & "|" & FootFontStrikethrough & "|" & FootFontColor
End Property

Public Function GetFormat(edtThis As Object) As String
    Dim strF As String
    With edtThis
        strF = .PaperKind                               '0
        strF = strF & ";" & .PaperOrient                '1
        strF = strF & ";" & .PaperHeight                '2
        strF = strF & ";" & .PaperWidth                 '3
        strF = strF & ";" & .MarginLeft                 '4
        strF = strF & ";" & .MarginRight                '5
        strF = strF & ";" & .MarginTop                  '6
        strF = strF & ";" & .MarginBottom               '7
        strF = strF & ";" & .BackColor                  '8
        strF = strF & ";" & .PaperColor                 '9
        strF = strF & ";" & IIf(.ShowPageNumber, 1, 0)  '10
        strF = strF & ";" & .HeadFontName & "|" & .HeadFontSize & "|" & .HeadFontBold & "|" & .HeadFontItalic & "|" & .HeadFontUnderline & "|" & .HeadFontStrikethrough & "|" & .HeadFontColor
        strF = strF & ";" & .FootFontName & "|" & .FootFontSize & "|" & .FootFontBold & "|" & .FootFontItalic & "|" & .FootFontUnderline & "|" & .FootFontStrikethrough & "|" & .FootFontColor
    End With
    GetFormat = strF
End Function

Public Sub SetFormat(edtThis As Object, strFmt As String)
    Dim T As Variant, T2 As Variant
    T = Split(strFmt, ";")
    On Error Resume Next
    With edtThis
        .PaperHeight = T(2)
        .PaperWidth = T(3)
        If T(0) > 0 Then
            .PaperKind = T(0)
        End If
        .PaperOrient = T(1)
        .MarginLeft = T(4)
        .MarginRight = T(5)
        .MarginTop = T(6)
        .MarginBottom = T(7)
        .BackColor = T(8)
        .PaperColor = T(9)
        .ShowPageNumber = (T(10) = 1)
        T2 = Split(T(11), "|")
        .HeadFontName = T2(0)
        .HeadFontSize = T2(1)
        .HeadFontBold = T2(2)
        .HeadFontItalic = T2(3)
        .HeadFontUnderline = T2(4)
        .HeadFontStrikethrough = T2(5)
        .HeadFontColor = T2(6)
        T2 = Split(T(12), "|")
        .FootFontName = T2(0)
        .FootFontSize = T2(1)
        .FootFontBold = T2(2)
        .HeadFontItalic = T2(3)
        .FootFontUnderline = T2(4)
        .HeadFontStrikethrough = T2(5)
        .FootFontColor = T2(6)
    End With
End Sub
Public Function ReadPagePicture() As StdPicture
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  intKind-病历种类; strPageNo-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strPicFile As String
    Set ReadPagePicture = Nothing
    strPicFile = zlBlobRead(7, Me.种类 & "-" & Me.页面)
    If gobjFSO.FileExists(strPicFile) Then
        Err = 0: On Error Resume Next
        Set ReadPagePicture = LoadPicture(strPicFile)
        gobjFSO.DeleteFile strPicFile, True      '删除临时文件
    End If
End Function
Public Sub ReadPageHead(edtThis As Object)
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  intKind-病历种类; strPageNo-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String
    strFile = zlBlobRead(12, Me.种类 & "-" & Me.页面, App.Path & "\Head.rtf")
    If gobjFSO.FileExists(strFile) Then
        edtThis.HeadFile = strFile           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        If Trim(edtThis.HeadFileText) = "" Then GoTo Headtxt
    Else
Headtxt:
        edtThis.Head = Me.页眉
        edtThis.HeadTextToFile '将文字读入Rtf控件中
    End If
End Sub
Public Sub ReadPageFoot(edtThis As Object)
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  intKind-病历种类; strPageNo-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strFoot As String
    strFile = zlBlobRead(13, Me.种类 & "-" & Me.页面, App.Path & "\Foot.rtf")
    If gobjFSO.FileExists(strFile) Then
        edtThis.FootFile = strFile            '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        If Trim(edtThis.FootFileText) = "" Then GoTo Foottxt
    Else
Foottxt:
        edtThis.Foot = Me.页脚
        edtThis.FootTextToFile '将文字读入Rtf控件中
    End If
End Sub
