VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'######################################################################################################################

Public Enum CellProperty
    AligmentV = 1                           '垂直对齐
    AligmentH = 2                           '水平对齐
    FontName = 3
    FontSize = 4
    ForeColor = 5
    WordWarp = 6
    MuliLine = 7
    AutoFit = 8
    Data = 9
    OutPutV = 10
End Enum

Private mrsGridFormat As ADODB.Recordset
Private mrsGridData As ADODB.Recordset
Private mrsGridRowFormat As ADODB.Recordset
Private mrsGridColFormat As ADODB.Recordset

Private mclsCommonLib As clsCommonLib

'######################################################################################################################
Public Function InitGrid(ByVal bytAligment As Byte, ByVal lngForeColor As Long, ByVal lngBackColor As Long, ByVal strFontName As String, ByVal sglFontNameSize As Single, Optional ByVal strGridLine As String = "111111", Optional ByVal bytAligmentV As Byte = 2) As Boolean
    '******************************************************************************************************************
    '功能： 初始表格管理数据
    '参数： rsData　- 单元格数据；rsRowFormat -  行格式说明；rsColFormat - 列格式说明；
    '返回：
    '******************************************************************************************************************
    
    Set mrsGridFormat = New ADODB.Recordset
    With mrsGridFormat
        .Fields.Append "表格线", adVarChar, 12
        .Fields.Append "前景色", adVarChar, 20
        .Fields.Append "背景色", adVarChar, 20
        .Fields.Append "字体", adVarChar, 20
        .Fields.Append "大小", adVarChar, 10
        .Fields.Append "水平对齐", adVarChar, 2
        .Fields.Append "垂直对齐", adVarChar, 2
        .Open
    End With
        
    Set mrsGridColFormat = New ADODB.Recordset
    With mrsGridColFormat
        .Fields.Append "列号", adTinyInt
        .Fields.Append "列宽", adVarChar, 30
        .Fields.Append "线型", adTinyInt
        .Fields.Append "线宽", adTinyInt
        .Open
    End With
    
    Set mrsGridRowFormat = New ADODB.Recordset
    With mrsGridRowFormat
        .Fields.Append "行号", adTinyInt
        .Fields.Append "行高", adVarChar, 30
        .Fields.Append "线型", adTinyInt
        .Fields.Append "线宽", adTinyInt
        .Open
    End With
    
    Set mrsGridData = New ADODB.Recordset
    With mrsGridData
        .Fields.Append "行号", adTinyInt
        .Fields.Append "列号", adTinyInt
        .Fields.Append "内容", adVarChar, 1000
        .Fields.Append "前景色", adVarChar, 20
        .Fields.Append "背景色", adVarChar, 20
        .Fields.Append "字体", adVarChar, 20
        .Fields.Append "大小", adVarChar, 10
        .Fields.Append "水平对齐", adVarChar, 2
        .Fields.Append "垂直对齐", adVarChar, 2
        .Fields.Append "自动换行", adTinyInt
        .Fields.Append "多行输出", adTinyInt
        .Fields.Append "自动适应", adTinyInt
        .Fields.Append "旋转角度", adTinyInt   '0-正常;1-从上到下纵排;2-从上到下纵排左转90度;3-从下到上纵排;4-从下到上右转90度
        .Open
    End With

    mrsGridFormat.AddNew
    mrsGridFormat("水平对齐").value = bytAligment
    mrsGridFormat("垂直对齐").value = bytAligmentV
    mrsGridFormat("前景色").value = lngForeColor
    mrsGridFormat("背景色").value = lngBackColor
    mrsGridFormat("字体").value = strFontName
    mrsGridFormat("大小").value = sglFontNameSize
    mrsGridFormat("表格线").value = Mid(strGridLine, 1, 1) & ";" & Mid(strGridLine, 2, 1) & ";" & Mid(strGridLine, 3, 1) & ";" & Mid(strGridLine, 4, 1) & ";" & Mid(strGridLine, 5, 1) & ";" & Mid(strGridLine, 6, 1)
    
    InitGrid = True
    
End Function

Public Property Let Cell(ByVal Property As CellProperty, ByVal intRow As Integer, ByVal intCol As Integer, ByVal vData As Variant)

    mrsGridData.Filter = ""
    mrsGridData.Filter = "行号=" & intRow & " And 列号=" & intCol
    
    If mrsGridData.RecordCount > 0 Then
        Select Case Property
        Case CellProperty.AligmentH
            mrsGridData("水平对齐").value = Val(vData)
        Case CellProperty.AligmentV
            mrsGridData("垂直对齐").value = Val(vData)
        Case CellProperty.FontName
            mrsGridData("字体").value = Trim(vData)
        Case CellProperty.FontSize
            mrsGridData("大小").value = Val(vData)
        Case CellProperty.ForeColor
            mrsGridData("前景色").value = Val(vData)
        Case CellProperty.WordWarp
            mrsGridData("自动换行").value = Val(vData)
        Case CellProperty.MuliLine
            mrsGridData("多行输出").value = Val(vData)
        Case CellProperty.AutoFit
            mrsGridData("自动适应").value = Val(vData)
        Case CellProperty.Data
            mrsGridData("内容").value = vData
        Case CellProperty.OutPutV
            mrsGridData("旋转角度").value = Val(vData)
        End Select
    Else
        Select Case Property
        Case CellProperty.Data
    
            mrsGridData.AddNew
            mrsGridData("行号").value = intRow
            mrsGridData("列号").value = intCol
            mrsGridData("内容").value = vData
            mrsGridData("字体").value = mrsGridFormat("字体").value
            mrsGridData("大小").value = mrsGridFormat("大小").value
            mrsGridData("前景色").value = mrsGridFormat("前景色").value
            mrsGridData("背景色").value = mrsGridFormat("背景色").value
            mrsGridData("水平对齐").value = mrsGridFormat("水平对齐").value
            mrsGridData("垂直对齐").value = mrsGridFormat("垂直对齐").value
            mrsGridData("自动换行").value = 0
            mrsGridData("多行输出").value = 0
            mrsGridData("自动适应").value = 0
            mrsGridData("旋转角度").value = 0
        End Select
    End If
    

        
End Property

Public Property Let RowCell(ByVal Property As CellProperty, ByVal intRow1 As Integer, ByVal intRow2 As Integer, ByVal vData As Variant)

    mrsGridData.Filter = ""
    mrsGridData.Filter = "行号>=" & intRow1 & " And 行号<=" & intRow2
    
    If mrsGridData.RecordCount > 0 Then
        mrsGridData.MoveFirst
        Do While Not mrsGridData.EOF
        
            Select Case Property
            Case CellProperty.AligmentH
                mrsGridData("水平对齐").value = Val(vData)
            Case CellProperty.AligmentV
                mrsGridData("垂直对齐").value = Val(vData)
            Case CellProperty.FontName
                mrsGridData("字体").value = Trim(vData)
            Case CellProperty.FontSize
                mrsGridData("大小").value = Val(vData)
            Case CellProperty.ForeColor
                mrsGridData("前景色").value = Val(vData)
            End Select
            
            mrsGridData.MoveNext
        Loop
    End If
    
End Property

Public Property Let ColCell(ByVal Property As CellProperty, ByVal intCol1 As Integer, ByVal intCol2 As Integer, ByVal vData As Variant)

    mrsGridData.Filter = ""
    mrsGridData.Filter = "列号>=" & intCol1 & " And 列号<=" & intCol2
    
    If mrsGridData.RecordCount > 0 Then
        mrsGridData.MoveFirst
        Do While Not mrsGridData.EOF
        
            Select Case Property
            Case CellProperty.AligmentH
                mrsGridData("水平对齐").value = Val(vData)
            Case CellProperty.AligmentV
                mrsGridData("垂直对齐").value = Val(vData)
            Case CellProperty.FontName
                mrsGridData("字体").value = Trim(vData)
            Case CellProperty.FontSize
                mrsGridData("大小").value = Val(vData)
            Case CellProperty.ForeColor
                mrsGridData("前景色").value = Val(vData)
            End Select
            
            mrsGridData.MoveNext
        Loop
    End If
    
End Property


Public Function SetGridRowFormat(ByVal strRowHeight As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strRowHeight, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGridRowFormat.AddNew
        mrsGridRowFormat("行号").value = intLoop + 1
        mrsGridRowFormat("行高").value = mclsCommonLib.GetTwipsY(Val(varTmp(intLoop)))
        mrsGridRowFormat("线宽").value = 1
        mrsGridRowFormat("线型").value = 0
    Next
    
    SetGridRowFormat = True
End Function

Public Function SetGridRowHeight(ByVal intRow As Integer, ByVal sglHeight As Single) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mrsGridRowFormat.Filter = ""
    mrsGridRowFormat.Filter = "行号=" & intRow + 1
    If mrsGridRowFormat.RecordCount > 0 Then
        mrsGridRowFormat("行高").value = mclsCommonLib.GetTwipsY(sglHeight)
    End If
    mrsGridRowFormat.Filter = ""

    SetGridRowHeight = True
    
End Function

Public Function SetGridRowLineType(ByVal intRow As Integer, ByVal bytLineType As Byte) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mrsGridRowFormat.Filter = ""
    mrsGridRowFormat.Filter = "行号=" & intRow
    If mrsGridRowFormat.RecordCount > 0 Then
        mrsGridRowFormat("线宽").value = bytLineType
    End If
    
    SetGridRowLineType = True
End Function

Public Function SetGridRowLineStyle(ByVal intRow As Integer, ByVal bytLineStyle As Byte) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mrsGridRowFormat.Filter = ""
    mrsGridRowFormat.Filter = "行号=" & intRow
    If mrsGridRowFormat.RecordCount > 0 Then
        mrsGridRowFormat("线型").value = bytLineStyle
    End If
    
    SetGridRowLineStyle = True
End Function

Public Function SetGridColFormat(ByVal strColWidth As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strColWidth, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGridColFormat.AddNew
        mrsGridColFormat("列号").value = intLoop + 1
        mrsGridColFormat("列宽").value = mclsCommonLib.GetTwipsY(Val(varTmp(intLoop)))
        mrsGridColFormat("线宽").value = 1
        mrsGridRowFormat("线型").value = 0
    Next
    
    SetGridColFormat = True
    
End Function

Public Function SetGridColLineType(ByVal intCol As Integer, ByVal bytLineType As Byte) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mrsGridColFormat.Filter = ""
    mrsGridColFormat.Filter = "列号=" & intCol
    If mrsGridColFormat.RecordCount > 0 Then
        mrsGridColFormat("线宽").value = bytLineType
    End If
    
    SetGridColLineType = True
End Function


Public Function Width() As Single
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglWidth As Single
    Dim intCols As Integer
    
    sglWidth = 0
    mrsGridColFormat.Filter = ""
    intCols = mrsGridColFormat.RecordCount
    For intLoop = 1 To intCols

        mrsGridColFormat.Filter = ""
        mrsGridColFormat.Filter = "列号=" & intLoop
        If mrsGridColFormat.RecordCount > 0 Then
'
            sglWidth = sglWidth + Val(mrsGridColFormat("列宽").value)

        End If
    Next
    
    Width = mclsCommonLib.GetCentimetersX(sglWidth)
End Function

Public Function SumGridColWidth(Optional ByVal intCol As Integer) As Single
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglWidth As Single
    Dim intCols As Integer
    
    sglWidth = 0
    mrsGridColFormat.Filter = ""
    intCols = mrsGridColFormat.RecordCount
    For intLoop = 1 To intCols

        mrsGridColFormat.Filter = ""
        mrsGridColFormat.Filter = "列号=" & intLoop
        If mrsGridColFormat.RecordCount > 0 Then
'
            sglWidth = sglWidth + Val(mrsGridColFormat("列宽").value)
            
            If intLoop = intCol Then
                SumGridColWidth = sglWidth
                Exit Function
            End If

        End If
    Next
    
    SumGridColWidth = sglWidth
End Function

Public Function Height(Optional ByVal intRow As Integer) As Single
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglHeight As Single
    Dim intRows As Integer
    
    sglHeight = 0
    
    mrsGridRowFormat.Filter = ""
    intRows = mrsGridRowFormat.RecordCount
    For intLoop = 1 To intRows

        mrsGridRowFormat.Filter = ""
        mrsGridRowFormat.Filter = "行号=" & intLoop
        If mrsGridRowFormat.RecordCount > 0 Then
            
            sglHeight = sglHeight + mclsCommonLib.GetCentimetersY(Val(mrsGridRowFormat("行高").value))
            
            If intLoop = intRow Then
                Height = sglHeight
                Exit Function
            End If
            
        End If
    Next
    
    Height = sglHeight
End Function

Public Function SumGridRowHeight(Optional ByVal intRow As Integer) As Single
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglHeight As Single
    Dim intRows As Integer
    
    sglHeight = 0
    
    mrsGridRowFormat.Filter = ""
    intRows = mrsGridRowFormat.RecordCount
    For intLoop = 1 To intRows

        mrsGridRowFormat.Filter = ""
        mrsGridRowFormat.Filter = "行号=" & intLoop
        If mrsGridRowFormat.RecordCount > 0 Then
'
            sglHeight = sglHeight + Val(mrsGridRowFormat("行高").value)
            
            If intLoop = intRow Then
                SumGridRowHeight = sglHeight
                Exit Function
            End If

        End If
    Next
    
    SumGridRowHeight = sglHeight
End Function

Public Function SetGridColData(ByVal intCol As Integer, ByVal strData As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strData, ";")
    
    
    For intLoop = 0 To UBound(varTmp)
        mrsGridData.AddNew
        mrsGridData("行号").value = intLoop + 1
        mrsGridData("列号").value = intCol
        mrsGridData("内容").value = CStr(varTmp(intLoop))
        mrsGridData("字体").value = mrsGridFormat("字体").value
        mrsGridData("大小").value = mrsGridFormat("大小").value
        mrsGridData("前景色").value = mrsGridFormat("前景色").value
        mrsGridData("背景色").value = mrsGridFormat("背景色").value
        mrsGridData("水平对齐").value = mrsGridFormat("水平对齐").value
        mrsGridData("垂直对齐").value = mrsGridFormat("垂直对齐").value
        mrsGridData("自动换行").value = 0
        mrsGridData("多行输出").value = 0
        mrsGridData("自动适应").value = 0
        mrsGridData("旋转角度").value = 0
'        mrsGridData("图形文件").Value = ""
    Next
    
    SetGridColData = True
    
End Function

Public Function SetGridRowData(ByVal intRow As Integer, ByVal strData As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strData, "'")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGridData.AddNew
        mrsGridData("行号").value = intRow
        mrsGridData("列号").value = intLoop + 1
        mrsGridData("内容").value = CStr(varTmp(intLoop))
        mrsGridData("字体").value = mrsGridFormat("字体").value
        mrsGridData("大小").value = mrsGridFormat("大小").value
        mrsGridData("前景色").value = mrsGridFormat("前景色").value
        mrsGridData("背景色").value = mrsGridFormat("背景色").value
        mrsGridData("水平对齐").value = mrsGridFormat("水平对齐").value
        mrsGridData("垂直对齐").value = mrsGridFormat("垂直对齐").value
        mrsGridData("自动换行").value = 0
        mrsGridData("多行输出").value = 0
        mrsGridData("自动适应").value = 0
        mrsGridData("旋转角度").value = 0
'        mrsGridData("图形文件").Value = ""
    Next
    
    SetGridRowData = True
    
End Function

Friend Function FormatData(Optional ByVal strKind As String = "数据") As Boolean
    '******************************************************************************************************************
    '功能： 产生表格数据
    '参数： mrsGridData　- 单元格数据；mrsGridRowFormat -  行格式说明；mrsGridColFormat - 列格式说明；
    '       gobjRect - 表格输出的左上角坐标和右下角坐标
    '返回：
    '******************************************************************************************************************
    Dim strTmp As String
    Dim intRows As Integer
    Dim intCols As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim X0 As Single
    Dim Y0 As Single
    Dim X1 As Single
    Dim Y1 As Single
    Dim varGridLine As Variant
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim objSvrFont As USERFONT
    Dim objSvrRect As USERRECT
    Dim sngOldY0 As Single
    
    On Error GoTo errHand
    
    Call SaveFont(gobjFont, objSvrFont)
    Call SaveRect(gobjRect, objSvrRect)
    
    X0 = gobjRect.X0
    Y0 = gobjRect.Y0
    X1 = gobjRect.X1
    Y1 = gobjRect.Y1
    
    mrsGridRowFormat.Filter = ""
    mrsGridColFormat.Filter = ""
    
    intRows = mrsGridRowFormat.RecordCount
    intCols = mrsGridColFormat.RecordCount
    
    '先画表格线
    '------------------------------------------------------------------------------------------------------------------
    varGridLine = Split(mrsGridFormat("表格线").value, ";")
    
    '外框：上横线
    If Val(varGridLine(0)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.X1 = X1
        gobjRect.Y1 = Y0
        Call AppendPrintData(strKind, "线条")
    End If
    
    '外框：右竖线
    If Val(varGridLine(1)) > 0 Then
        gobjRect.X0 = X1
        gobjRect.Y0 = Y0
        gobjRect.X1 = X1
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "线条")
    
    End If
    
    '外框：下横线
    If Val(varGridLine(2)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y1
        gobjRect.X1 = X1
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "线条")
    
    End If
    
    '外框：左竖线
    If Val(varGridLine(3)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.X1 = X0
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "线条")
            
    End If
    
    
    '内框：横线
    gobjFont.LineWidth = 1
    If Val(varGridLine(4)) > 0 Then
        gobjRect.Y0 = Y0
        gobjRect.X0 = X0
        gobjRect.X1 = X1
        For intRow = 1 To intRows - 1
            mrsGridRowFormat.Filter = ""
            mrsGridRowFormat.Filter = "行号=" & intRow
            If mrsGridRowFormat.RecordCount > 0 Then
                gobjRect.Y0 = gobjRect.Y0 + Val(mrsGridRowFormat("行高").value)
                gobjRect.Y1 = gobjRect.Y0
                If Val(mrsGridRowFormat("线宽").value) <> 0 Then
                    gobjFont.LineWidth = Val(mrsGridRowFormat("线宽").value)
                    gobjFont.LineStyle = Val(mrsGridRowFormat("线型").value)
                    Call AppendPrintData(strKind, "线条")
                End If
            End If
        Next
    End If
        
    '内框：竖线
    If Val(varGridLine(5)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.Y1 = Y1
        For intCol = 1 To intCols - 1
            mrsGridColFormat.Filter = ""
            mrsGridColFormat.Filter = "列号=" & intCol
            If mrsGridColFormat.RecordCount > 0 Then
                gobjRect.X0 = gobjRect.X0 + Val(mrsGridColFormat("列宽").value)
                gobjRect.X1 = gobjRect.X0
                gobjFont.LineWidth = Val(mrsGridColFormat("线宽").value)
                gobjFont.LineStyle = Val(mrsGridRowFormat("线型").value)
                Call AppendPrintData(strKind, "线条")
            End If
        Next
    End If
    gobjFont.LineWidth = 1
    
    '填写数据
    '------------------------------------------------------------------------------------------------------------------
    gobjRect.Y1 = Y0
    For intRow = 1 To intRows
        
        gobjRect.Y0 = gobjRect.Y1
        gobjRect.X0 = X0
        
        mrsGridRowFormat.Filter = ""
        mrsGridRowFormat.Filter = "行号=" & intRow
        
        If mrsGridRowFormat.RecordCount > 0 Then

            gobjRect.Y1 = gobjRect.Y0 + Val(mrsGridRowFormat("行高").value)
            gobjRect.X1 = X0
            
            For intCol = 1 To intCols
                
                mrsGridColFormat.Filter = ""
                mrsGridColFormat.Filter = "列号=" & intCol
                    
                If mrsGridColFormat.RecordCount > 0 Then
                
                    gobjRect.X0 = gobjRect.X1
                    gobjRect.X1 = gobjRect.X0 + Val(mrsGridColFormat("列宽").value)
                                        
                    mrsGridData.Filter = ""
                    mrsGridData.Filter = "行号=" & intRow & " And 列号=" & intCol
                    If mrsGridData.RecordCount > 0 Then
     
                        strTmp = mrsGridData("内容").value
                    
                        gobjFont.Name = mrsGridData("字体").value
                        gobjFont.size = mrsGridData("大小").value
                        gobjFont.BackColor = Val(mrsGridData("背景色").value)
                        gobjFont.ForeColor = Val(mrsGridData("前景色").value)

'                        gobjFont.Bold = False
                        gobjFont.Underline = False
                        Call SetDraw(gobjFont)
                        
                        '<Image>c:\a<Image>"
                        intPos1 = InStr(strTmp, "<Image>")
                        intPos2 = InStr(strTmp, "</Image>")
                        If intPos1 > 0 And intPos2 > 0 And intPos1 < intPos2 Then
                            strTmp = Mid(strTmp, intPos1 + 7, intPos2 - intPos1 - 7)
                            If Dir(strTmp) <> "" And strTmp <> "" Then
                                Call AppendPrintData(strKind, "图形", Val(mrsGridData("水平对齐").value), False, strTmp)
                            End If
                        Else
                        
                            If Val(mrsGridData("水平对齐").value) = 1 Then gobjRect.X0 = gobjRect.X0 + mclsCommonLib.GetTwipsX(0.05)
                            If Val(mrsGridData("垂直对齐").value) = 1 Then gobjRect.Y0 = gobjRect.Y0 + mclsCommonLib.GetTwipsY(0.05)
                        
                            If (Val(mrsGridData("多行输出").value) = 1) = True Then sngOldY0 = gobjRect.Y0
                            
                            Call AppendPrintData(strKind, "文本", Val(mrsGridData("水平对齐").value), (Val(mrsGridData("自动换行").value) = 1), strTmp, Val(mrsGridData("垂直对齐").value), (Val(mrsGridData("多行输出").value) = 1), gobjFont.Rows, (Val(mrsGridData("自动适应").value) = 1), , , Val(mrsGridData("旋转角度").value))
                            
                            If (Val(mrsGridData("多行输出").value) = 1) = True Then gobjRect.Y0 = sngOldY0
                        
                            If Val(mrsGridData("水平对齐").value) = 1 Then gobjRect.X0 = gobjRect.X0 - mclsCommonLib.GetTwipsX(0.05)
                            If Val(mrsGridData("垂直对齐").value) = 1 Then gobjRect.Y0 = gobjRect.Y0 - mclsCommonLib.GetTwipsY(0.05)
                        
                        End If

                        
                    End If
                End If
            Next
        End If
    Next
    
    Call SaveFont(objSvrFont, gobjFont)
    Call SaveRect(objSvrRect, gobjRect)
    
    FormatData = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveFont(objSvrFont, gobjFont)
End Function

Private Sub Class_Initialize()
    Set mclsCommonLib = New clsCommonLib
End Sub

Private Sub Class_Terminate()
    Set mclsCommonLib = Nothing
End Sub
