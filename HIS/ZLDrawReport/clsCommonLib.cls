VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'定义
'######################################################################################################################

Public Enum USERCOLOR
    兰色 = 16711680
    白色 = 16777215
    黑色 = 0
    背景色 = 10191921
    表格线色 = 5880731
    表格背景色 = 15790320

    项目标题背景色 = -2147483636
    
End Enum

Public Type USERRECT
    X0 As Single
    Y0 As Single
    X1 As Single
    Y1 As Single
    R0 As Single                          '行间距
    B0 As Single                          '边界
    Page As Integer                     '页号
End Type

Public Type USERFONT
    Name As String
    Size As Single
    Bold As Boolean
    Italic As Boolean
    Underline As Boolean
    Strikethrough As Boolean
    ForeColor As Long
    BackColor As Long
    LineWidth As Integer
    LineStyle As Integer
    HAlignment As Byte                              '横向对齐
    VAlignment As Byte                              '纵向对齐
    WordWarp As Boolean                             '自动换行
    MuliLine As Boolean                             '多行文本
    AutoFit As Boolean                              '自动适应
    Rows As Integer                                 '行数
End Type

Public Type USERPAPER
    Width As Single
    Height As Single
    BorderLeft As Single
    BorderRight As Single
    BorderTop As Single
    BorderBottom As Single
    PageHead As Single
    PageFoot As Single
    SpaceTop As Single
    SpaceBottom As Single
    PaperSize As Integer
    Orientation As Byte
End Type

Public Event FormatData()

'######################################################################################################################
Public Function InitOracle(cnOracle As ADODB.Connection)
    
    Set gcnOracle = cnOracle
    
End Function

Public Function SaveLob(ByVal lngSys As Long, ByVal lngAction As Long, ByVal strKeyWord As String, ByVal strFile As String, Optional ByVal bytFunc As Byte = 0) As Boolean
    SaveLob = sys.SaveLob(lngSys, lngAction, strKeyWord, strFile, bytFunc)
End Function

Public Function ReadLob(ByVal lngSys As Long, ByVal lngAction As Long, ByVal strKeyWord As String, _
                        Optional ByVal strFile As String, Optional ByVal bytFunc As Byte = 0, _
                        Optional bytMoved As Byte = 0) As String
    ReadLob = sys.ReadLob(lngSys, lngAction, strKeyWord, strFile, bytFunc, bytMoved)
End Function

Public Function ReadLobV2(ByVal strTable As String, _
                          ByVal strLOBField As String, _
                          ByVal strWhere As String, _
                          ByVal strFile As String, _
                          ParamArray arrInput() As Variant) As String
    ReadLobV2 = sys.ReadLobV2(strTable, strLOBField, strWhere, strFile, arrInput())
End Function

Public Property Get Rect() As USERRECT
    Rect = gobjRect
End Property

Public Property Let Rect(vData As USERRECT)
    gobjRect = vData
End Property

Public Property Get Font() As USERFONT
    Font = gobjFont
End Property

Public Property Let Font(vData As USERFONT)
    gobjFont = vData
    
    gobjDraw.FontName = gobjFont.Name
    gobjDraw.FontSize = gobjFont.Size
    gobjDraw.FontBold = gobjFont.Bold
    gobjDraw.FontItalic = gobjFont.Italic
    
End Property

Public Property Get Paper() As USERPAPER
    Paper = gobjPaper
End Property

Public Property Let Paper(vData As USERPAPER)
    gobjPaper = vData
End Property

Public Property Get Draw() As Object
    Set Draw = gobjDraw
End Property

Public Property Get DataSource(vntIndexKey As String) As clsDataSource
    Set DataSource = gclsDataSources.Item(vntIndexKey)
End Property

Public Property Get RegInfo(ByVal strItem As String) As String
    RegInfo = zlRegInfo(strItem)
End Property

Public Function GetTextWidth(ByVal strText As String) As Single
    GetTextWidth = gobjDraw.TextWidth(strText)
End Function

Public Function GetTextHeight(ByVal strText As String) As Single
    GetTextHeight = gobjDraw.TextHeight(strText)
End Function

Public Function GetTwipsX(ByVal sglNumber As Single) As Single
    GetTwipsX = gobjDraw.ScaleX(sglNumber, vbCentimeters, vbTwips)
End Function

Public Function GetTwipsY(ByVal sglNumber As Single) As Single
    GetTwipsY = gobjDraw.ScaleY(sglNumber, vbCentimeters, vbTwips)
End Function

Public Function GetCentimetersX(ByVal sglNumber As Single) As Single
    GetCentimetersX = gobjDraw.ScaleX(sglNumber, vbTwips, vbCentimeters)
End Function

Public Function GetCentimetersY(ByVal sglNumber As Single) As Single
    GetCentimetersY = gobjDraw.ScaleY(sglNumber, vbTwips, vbCentimeters)
End Function

Public Function GetRows(ByVal objDraw As Object, ByVal strText As String, ByVal lngCX As Long) As Integer
    '******************************************************************************************************************
    '功能：获取需要的行数，因为有可能要换行
    '参数：
    '返回：
    '******************************************************************************************************************
    GetRows = GetLines(objDraw, strText, lngCX)
End Function

Public Function ShowError() As Byte
    
    ShowError = ErrCenter
    
End Function

Public Sub SaveError()
    Call SaveErrLog
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    NVL = zlStr.NVL(varValue, DefaultValue)
End Function

Public Function PrintSetForm() As Object
    Set PrintSetForm = frmChildPrintSet
End Function

Public Function CreateParameter() As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Fields.Append "参数名称", adVarChar, 30
        .Fields.Append "参数结果", adVarChar, 4000
        .Fields.Append "参数类型", adVarChar, 30
        .Open
    End With
    
    Set CreateParameter = rs
    
End Function

Public Function SetParameter(ByRef rs As ADODB.Recordset, ByVal strParameterName As String, ByVal strParameterValue As String, Optional ByVal strParameterType As String = "文本") As Boolean
    
    rs.Filter = ""
    rs.Filter = "参数名称='" & strParameterName & "'"
    If rs.RecordCount = 0 Then rs.AddNew
    rs("参数名称").value = strParameterName
    rs("参数结果").value = strParameterValue
    rs("参数类型").value = strParameterType
    
    SetParameter = True
    
End Function

Public Function GetParameter(ByRef rs As ADODB.Recordset, ByVal strParameterName As String) As String
    
    rs.Filter = ""
    rs.Filter = "参数名称='" & strParameterName & "'"
    If rs.RecordCount > 0 Then
        GetParameter = CStr(rs("参数结果").value)
    End If
    
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
    '       实际数据存储长度
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function GetListRecord() As ADODB.Recordset
    If grsList.RecordCount > 0 Then grsList.MoveFirst
    Set GetListRecord = grsList
End Function

Public Function DatabaseObject() As clsDatabase
    
    Set DatabaseObject = zlDatabase
    
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant, i As Long
    arrPars = arrInput
    Set OpenSQLRecord = zlDatabase.OpenSQLRecordByArray(strSQL, strTitle, arrPars)
    
End Function

Public Function FixText(ByVal strPreText As String, ByVal strLstText As String, ByVal lngCX As Single, ByVal strRepaceChar As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strReplace As String
    Dim strText As String
    
    strText = strPreText & strLstText

    gobjDraw.FontName = gobjFont.Name
    gobjDraw.FontSize = gobjFont.Size
    gobjDraw.FontBold = gobjFont.Bold
    gobjDraw.FontItalic = gobjFont.Italic
        
    Do While (True)
        strReplace = strReplace & strRepaceChar
        If gobjDraw.TextWidth(strText & strReplace & strRepaceChar) > lngCX Then
            Exit Do
        End If
    Loop
    
    FixText = strPreText & strReplace & strLstText
    
End Function

Public Function ReadPicture(rsTable As ADODB.Recordset, strField As String, Optional strFile As String) As String
    ReadPicture = zlDatabase.ReadPicture(rsTable, strField, strFile)
End Function

Public Function GetRowText(ByVal objDraw As Object, ByVal strText As String, ByVal intRow As Integer, ByVal lngCX As Long) As String
    '******************************************************************************************************************
    '功能：获取指定行的数据，方法是先求出可以最多输出多少个字符，然后调用过程“Get_LineText”求出指定行内容
    '参数：
    '返回：
    '******************************************************************************************************************

    GetRowText = GetLineText2(objDraw, strText, intRow, lngCX)
End Function

Private Sub Class_Initialize()
    gobjFont.Name = IIf(gobjFont.Name = "", "宋体", gobjFont.Name)
    gobjFont.Size = IIf(gobjFont.Size = 0, 12, gobjFont.Size)
    gobjFont.Bold = False
    gobjFont.Underline = False
    gobjFont.Strikethrough = False
    gobjFont.Italic = False
End Sub
