VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabFileDefineInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'病历文件定义
Public ID As Long               'ID
Public 种类 As EPRDocType   '种类：1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-诊断文书;6-知情文件;7-诊疗报告;8-诊疗申请
Public 编号 As String             '文件编号
Public 名称 As String           '文件名称
Public 说明 As String           '文件说明
Public 页面 As String           '文件页面：只有住院病历和护理记录可以设置页面,其他类型的病历独立成页,页面编号和文件编号一致
Public 保留 As Long             '是否保留：0-用户定义的病历文件;1-系统保留的普通病历文件,不能删除;<0,指单独编制程序实现的特殊病历,不能删除
Public 通用 As Long             '通用性：0-未用;1-通用;2-专科
Public 页面名称 As String       '病历页面名称
Public 报表 As Long             '对诊疗单据:0-可编辑,并按编辑格式输出;1-无编辑,用自定义报表输出(如处方笺);2-可编辑,用自定义报表输出(如各类检查报告)；对护理记录:0-病人记录;1-病房记录。其他无效。

Public lngModule As Long        '模块号

'格式 As String           'PaperKind;PaperOrient;PaperHeight;PaperWidth;MarginTop;MarginBottom;MarginLeft;MarginRight;BackColor;PaperColor;ShowPageNumber
'                            ;页眉HeadFontName|HeadFontSize|HeadFontBold|HeadFontItalic|HeadFontUnderline|HeadFontStrikethrough|HeadFontColor|HeadMargin
'                            ;页脚FootFontName|FootFontSize|FootFontBold|FootFontItalic|FootFontUnderline|FootFontStrikethrough|FootFontColor|FootMargin
'                            ;HCenter;VCenter
Private mvarPaperKind As Integer
Private mvarPaperOrient As Integer
Private mvarPaperHeight As Long
Private mvarPaperWidth  As Long
Private mvarMarginLeft As Long
Private mvarMarginRight As Long
Private mvarMarginTop As Long
Private mvarMarginBottom As Long
Private mvarHCenter As Byte '横向居中
Private mvarVCenter As Byte '纵向居中

'页眉                   '
Private mvarHeadConText As String          '内容(要素用[]包括),页眉内容
Private mvarHeadFontName As String
Private mvarHeadFontSize As Long
Private mvarHeadFontBold As Boolean
Private mvarHeadFontItalic As Boolean
Private mvarHeadFontUnderline As Boolean
Private mvarHeadFontStrikethrough As Boolean
Private mvarHeadFontColor As OLE_COLOR
Private mvarHeadMargin As Long
Private mvarHeadPic As StdPicture

'页脚                  '
Private mvarFootConText As String               '内容(要素用[]包括),页脚内容
Private mvarFootFontName As String
Private mvarFootFontSize As Long
Private mvarFootFontBold As Boolean
Private mvarFootFontItalic As Boolean
Private mvarFootFontUnderline As Boolean
Private mvarFootFontStrikethrough As Boolean
Private mvarFootFontColor As OLE_COLOR
Private mvarFootMargin As Long
Public Sub GetFileDefineInfo(ByVal lngFileID As Long)
Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
    gstrSQL = "Select a.ID, a.种类, a.编号, a.名称, a.说明, a.页面, a.保留, a.通用, b.名称 As 页面名称, b.报表, b.格式, b.页眉, b.页脚 " & _
            " From 病历文件列表 a, 病历页面格式 b " & _
            " Where a.页面 = b.编号 And a.种类 = b.种类 And a.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cTableEPR", lngFileID)
    With Me
        .ID = lngFileID
        If Not rsTemp.EOF Then
            .种类 = NVL(rsTemp!种类, 1)  '1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-疾病证明报告;6-知情文件;7-诊疗报告;8-诊疗申请
            .编号 = NVL(rsTemp!编号)
            .名称 = NVL(rsTemp!名称)
            .说明 = NVL(rsTemp!说明)
            .页面 = NVL(rsTemp!页面)
            .保留 = NVL(rsTemp!保留, 0)  '保留定义2-表格式病历
            .通用 = NVL(rsTemp!通用, 0)  '除诊疗单据外有效:0-未用;1-通用;2-专科
            .报表 = NVL(rsTemp!报表, 0)
            .页面名称 = NVL(rsTemp!页面名称)
            .格式 = NVL(rsTemp!格式, "9;1;16840;11907;482;805;350;350;16777215;16777215;1;宋体|9|False|False|False|False|vbBlack|350;宋体|9|False|False|False|False|vbBlack|350;0;0") '缺省 A4 纵向 297 210
            .页眉 = NVL(rsTemp!页眉, "")
            .页脚 = NVL(rsTemp!页脚, "")
            
            Dim strFile As String, strZipFile As String
            strZipFile = zlBlobRead(7, Me.种类 & "-" & Me.页面, strZipFile)
            If strZipFile = "" Then Exit Sub
            If gobjFSO.FileExists(strZipFile) Then
                strFile = zlFileUnzip(strZipFile, "Jpg")
                gobjFSO.DeleteFile strZipFile, True
                If strFile = "" Then Exit Sub
                
                If gobjFSO.FileExists(strFile) Then
                    Set Me.HeadPic = LoadPicture(strFile)
                    gobjFSO.DeleteFile strFile, True
                End If
            End If
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Property Let PaperKind(vData As Integer)
    mvarPaperKind = vData
End Property
Public Property Get PaperKind() As Integer
    PaperKind = mvarPaperKind
End Property
Public Property Let PaperOrient(vData As Integer)
    mvarPaperOrient = vData
End Property
Public Property Get PaperOrient() As Integer
    PaperOrient = mvarPaperOrient
End Property

Public Property Let PaperHeight(vData As Long)
    mvarPaperHeight = vData
End Property
Public Property Get PaperHeight() As Long
    PaperHeight = mvarPaperHeight
End Property
Public Property Let PaperWidth(vData As Long)
    mvarPaperWidth = vData
End Property
Public Property Get PaperWidth() As Long
    PaperWidth = mvarPaperWidth
End Property
Public Property Let MarginTop(vData As Long)
    mvarMarginTop = vData
End Property
Public Property Get MarginTop() As Long
    MarginTop = mvarMarginTop
End Property
Public Property Let MarginBottom(vData As Long)
    mvarMarginBottom = vData
End Property
Public Property Get MarginBottom() As Long
    MarginBottom = mvarMarginBottom
End Property
Public Property Let MarginLeft(vData As Long)
    mvarMarginLeft = vData
End Property
Public Property Get MarginLeft() As Long
    MarginLeft = mvarMarginLeft
End Property
Public Property Let MarginRight(vData As Long)
    mvarMarginRight = vData
End Property
Public Property Get MarginRight() As Long
    MarginRight = mvarMarginRight
End Property
Public Property Let HCenter(vData As Byte)
    mvarHCenter = vData
End Property
Public Property Get HCenter() As Byte
    HCenter = mvarHCenter
End Property
Public Property Let VCenter(vData As Byte)
    mvarVCenter = vData
End Property
Public Property Get VCenter() As Byte
    VCenter = mvarVCenter
End Property
Public Property Let 格式(vData As String)
On Error Resume Next
    Dim T As Variant, THead As Variant, TFoot As Variant
    T = Split(vData, ";")
    mvarPaperKind = T(0)
    mvarPaperOrient = T(1)
    mvarPaperHeight = T(2)
    mvarPaperWidth = T(3)
    mvarMarginTop = T(4)
    mvarMarginBottom = T(5)
    mvarMarginLeft = T(6)
    mvarMarginRight = T(7)
    
    THead = Split(T(11), "|")
    mvarHeadFontName = UsableFont(THead(0))
    mvarHeadFontSize = THead(1)
    mvarHeadFontBold = THead(2)
    mvarHeadFontItalic = THead(3)
    mvarHeadFontUnderline = THead(4)
    mvarHeadFontStrikethrough = THead(5)
    mvarHeadFontColor = THead(6)
    mvarHeadMargin = THead(7)
    If mvarHeadMargin = 0 Then
        mvarHeadMargin = mvarMarginTop
    End If
    
    TFoot = Split(T(12), "|")
    mvarFootFontName = UsableFont(TFoot(0))
    mvarFootFontSize = TFoot(1)
    mvarFootFontBold = TFoot(2)
    mvarFootFontItalic = TFoot(3)
    mvarFootFontUnderline = TFoot(4)
    mvarFootFontStrikethrough = TFoot(5)
    mvarFootFontColor = TFoot(6)
    mvarFootMargin = TFoot(7)
    If mvarFootMargin = 0 Then
        mvarFootMargin = mvarMarginBottom
    End If
    mvarHCenter = T(13)
    mvarVCenter = T(14)
    Err.Clear
End Property
Public Property Get 格式() As String
    格式 = mvarPaperKind & ";" & mvarPaperOrient & ";" & mvarPaperHeight & ";" & mvarPaperWidth & ";" & mvarMarginTop & ";" & mvarMarginBottom & ";" & mvarMarginLeft & ";" & mvarMarginRight & ";16777215;16777215;1" & _
        ";" & mvarHeadFontName & "|" & mvarHeadFontSize & "|" & mvarHeadFontBold & "|" & mvarHeadFontItalic & "|" & mvarHeadFontUnderline & "|" & mvarHeadFontStrikethrough & "|" & mvarHeadFontColor & "|" & mvarHeadMargin & _
        ";" & mvarFootFontName & "|" & mvarFootFontSize & "|" & mvarFootFontBold & "|" & mvarFootFontItalic & "|" & mvarFootFontUnderline & "|" & mvarFootFontStrikethrough & "|" & mvarFootFontColor & "|" & mvarFootMargin & ";" & mvarHCenter & ";" & mvarVCenter
End Property
Public Property Get HeadConText() As String
    HeadConText = mvarHeadConText
End Property
Public Property Let HeadConText(vData As String)
    mvarHeadConText = vData
End Property
Public Property Get HeadFontName() As String
    HeadFontName = mvarHeadFontName
End Property
Public Property Let HeadFontName(vData As String)
    mvarHeadFontName = vData
End Property
Public Property Get HeadFontSize() As Long
    HeadFontSize = mvarHeadFontSize
End Property
Public Property Let HeadFontSize(vData As Long)
    mvarHeadFontSize = vData
End Property
Public Property Get HeadFontBold() As Boolean
    HeadFontBold = mvarHeadFontBold
End Property
Public Property Let HeadFontBold(vData As Boolean)
    mvarHeadFontBold = vData
End Property
Public Property Get HeadFontItalic() As Boolean
    HeadFontItalic = mvarHeadFontItalic
End Property
Public Property Let HeadFontItalic(vData As Boolean)
    mvarHeadFontItalic = vData
End Property
Public Property Get HeadFontUnderline() As Boolean
    HeadFontUnderline = mvarHeadFontUnderline
End Property
Public Property Let HeadFontUnderline(vData As Boolean)
    mvarHeadFontUnderline = vData
End Property
Public Property Get HeadFontStrikethrough() As Boolean
    HeadFontStrikethrough = mvarHeadFontStrikethrough
End Property
Public Property Let HeadFontStrikethrough(vData As Boolean)
    mvarHeadFontStrikethrough = vData
End Property
Public Property Get HeadFontColor() As OLE_COLOR
    HeadFontColor = mvarHeadFontColor
End Property
Public Property Let HeadFontColor(vData As OLE_COLOR)
    mvarHeadFontColor = vData
End Property
Public Property Get HeadMargin() As Long
    HeadMargin = mvarHeadMargin
End Property
Public Property Let HeadMargin(vData As Long)
    mvarHeadMargin = vData
    If mvarHeadMargin = 0 Then
        If mvarPaperOrient = 1 Then mvarHeadMargin = mvarMarginTop Else mvarHeadMargin = mvarMarginLeft
    End If
End Property
Public Property Get 页眉() As String
    页眉 = mvarHeadConText
End Property
Public Property Let 页眉(vData As String)
    mvarHeadConText = vData
End Property
Public Property Get FootConText() As String
    FootConText = mvarFootConText
End Property
Public Property Let FootConText(vData As String)
    mvarFootConText = vData
End Property
Public Property Get FootFontName() As String
    FootFontName = mvarFootFontName
End Property
Public Property Let FootFontName(vData As String)
    mvarFootFontName = vData
End Property
Public Property Get FootFontSize() As Long
    FootFontSize = mvarFootFontSize
End Property
Public Property Let FootFontSize(vData As Long)
    mvarFootFontSize = vData
End Property
Public Property Get FootFontBold() As Boolean
    FootFontBold = mvarFootFontBold
End Property
Public Property Let FootFontBold(vData As Boolean)
    mvarFootFontBold = vData
End Property
Public Property Get FootFontItalic() As Boolean
    FootFontItalic = mvarFootFontItalic
End Property
Public Property Let FootFontItalic(vData As Boolean)
    mvarFootFontItalic = vData
End Property
Public Property Get FootFontUnderline() As Boolean
    FootFontUnderline = mvarFootFontUnderline
End Property
Public Property Let FootFontUnderline(vData As Boolean)
    mvarFootFontUnderline = vData
End Property
Public Property Get FootFontStrikethrough() As Boolean
    FootFontStrikethrough = mvarFootFontStrikethrough
End Property
Public Property Let FootFontStrikethrough(vData As Boolean)
    mvarFootFontStrikethrough = vData
End Property
Public Property Get FootFontColor() As OLE_COLOR
    FootFontColor = mvarFootFontColor
End Property
Public Property Let FootFontColor(vData As OLE_COLOR)
    mvarFootFontColor = vData
End Property
Public Property Get FootMargin() As Long
    FootMargin = mvarFootMargin
End Property
Public Property Let FootMargin(vData As Long)
    mvarFootMargin = vData
    If mvarFootMargin = 0 Then
        If mvarPaperOrient = 1 Then mvarFootMargin = mvarMarginBottom Else mvarFootMargin = mvarMarginRight
    End If
End Property
Public Property Get 页脚() As String
    页脚 = mvarFootConText
End Property
Public Property Let 页脚(vData As String)
    mvarFootConText = vData
End Property
Public Property Get HeadPic() As StdPicture
    Set HeadPic = mvarHeadPic
End Property
Public Property Set HeadPic(ByVal vData As StdPicture)
    Set mvarHeadPic = vData
End Property
Public Property Let HeadPic(ByVal vData As StdPicture)
    Set mvarHeadPic = vData
End Property
Public Function GetSaveSQL(arrSQL As Variant) As Boolean
    On Error GoTo errHand
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_病历页面格式_Update(" & Me.种类 & ",'" & Me.编号 & "','" & Me.页面名称 & "'," & _
                                            Me.报表 & ",'" & Me.格式 & "','" & Me.页眉 & "','" & Me.页脚 & "')"
                                            
    If mvarHeadPic.Handle <> 0 Then '有页眉图
        Dim strFile As String, strZipFile As String
        strFile = App.Path & "\Tmp.Jpg"
        If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True
        SavePicture mvarHeadPic, strFile
        strZipFile = zlFileZip(strFile)
        If strZipFile <> "" Then gobjFSO.DeleteFile strFile, True
        If gobjFSO.FileExists(strZipFile) Then
            Call zlBlobSql(7, Me.种类 & "-" & Me.页面, strZipFile, arrSQL)
            gobjFSO.DeleteFile strZipFile, True      '删除临时文件
        End If
    End If

    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_病历文件结构_Commit(" & Me.ID & ")"
    
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Class_Initialize()
    Set mvarHeadPic = New StdPicture
     mvarHeadFontName = "宋体"
     mvarHeadFontSize = 9
     mvarFootFontName = "宋体"
     mvarFootFontSize = 9
End Sub

Private Sub Class_Terminate()
    Set mvarHeadPic = Nothing
End Sub
