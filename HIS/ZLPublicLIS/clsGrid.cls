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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private mrsGrIDFormat As ADODB.Recordset
Private mrsGrIDData As ADODB.Recordset
Private mrsGrIDRowFormat As ADODB.Recordset
Private mrsGrIDColFormat As ADODB.Recordset

Private mclsLisReportLib As clsLisReportLib

'######################################################################################################################
Public Function InitGrID(ByVal bytAligment As Byte, ByVal lngForeColor As Long, ByVal lngBackColor As Long, ByVal strFontName As String, ByVal sglFontNameSize As Single, Optional ByVal strGrIDLine As String = "111111", Optional ByVal bytAligmentV As Byte = 2) As Boolean
    '******************************************************************************************************************
    '功能： 初始表格管理数据
    '参数： rsData　- 单元格数据；rsRowFormat -  行格式说明；rsColFormat - 列格式说明；
    '返回：
    '******************************************************************************************************************
    
    Set mrsGrIDFormat = New ADODB.Recordset
    With mrsGrIDFormat
        .Fields.Append "表格线", adVarChar, 12
        .Fields.Append "前景色", adVarChar, 20
        .Fields.Append "背景色", adVarChar, 20
        .Fields.Append "字体", adVarChar, 20
        .Fields.Append "大小", adVarChar, 10
        .Fields.Append "水平对齐", adVarChar, 2
        .Fields.Append "垂直对齐", adVarChar, 2
        .Open
    End With
        
    Set mrsGrIDColFormat = New ADODB.Recordset
    With mrsGrIDColFormat
        .Fields.Append "列号", adTinyInt
        .Fields.Append "列宽", adVarChar, 30
        .Fields.Append "线型", adTinyInt
        .Fields.Append "线宽", adTinyInt
        .Open
    End With
    
    Set mrsGrIDRowFormat = New ADODB.Recordset
    With mrsGrIDRowFormat
        .Fields.Append "行号", adTinyInt
        .Fields.Append "行高", adVarChar, 30
        .Fields.Append "线型", adTinyInt
        .Fields.Append "线宽", adTinyInt
        .Open
    End With
    
    Set mrsGrIDData = New ADODB.Recordset
    With mrsGrIDData
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

    mrsGrIDFormat.AddNew
    mrsGrIDFormat("水平对齐").Value = bytAligment
    mrsGrIDFormat("垂直对齐").Value = bytAligmentV
    mrsGrIDFormat("前景色").Value = lngForeColor
    mrsGrIDFormat("背景色").Value = lngBackColor
    mrsGrIDFormat("字体").Value = strFontName
    mrsGrIDFormat("大小").Value = sglFontNameSize
    mrsGrIDFormat("表格线").Value = Mid(strGrIDLine, 1, 1) & ";" & Mid(strGrIDLine, 2, 1) & ";" & Mid(strGrIDLine, 3, 1) & ";" & Mid(strGrIDLine, 4, 1) & ";" & Mid(strGrIDLine, 5, 1) & ";" & Mid(strGrIDLine, 6, 1)
    
    InitGrID = True
    
End Function

Public Property Let Cell(ByVal Property As CellProperty, ByVal intRow As Integer, ByVal intCol As Integer, ByVal vData As Variant)

    mrsGrIDData.Filter = ""
    mrsGrIDData.Filter = "行号=" & intRow & " And 列号=" & intCol
    
    If mrsGrIDData.RecordCount > 0 Then
        Select Case Property
        Case CellProperty.AligmentH
            mrsGrIDData("水平对齐").Value = Val(vData)
        Case CellProperty.AligmentV
            mrsGrIDData("垂直对齐").Value = Val(vData)
        Case CellProperty.FontName
            mrsGrIDData("字体").Value = Trim(vData)
        Case CellProperty.FontSize
            mrsGrIDData("大小").Value = Val(vData)
        Case CellProperty.ForeColor
            mrsGrIDData("前景色").Value = Val(vData)
        Case CellProperty.WordWarp
            mrsGrIDData("自动换行").Value = Val(vData)
        Case CellProperty.MuliLine
            mrsGrIDData("多行输出").Value = Val(vData)
        Case CellProperty.AutoFit
            mrsGrIDData("自动适应").Value = Val(vData)
        Case CellProperty.Data
            mrsGrIDData("内容").Value = vData
        Case CellProperty.OutPutV
            mrsGrIDData("旋转角度").Value = Val(vData)
        End Select
    Else
        Select Case Property
        Case CellProperty.Data
    
            mrsGrIDData.AddNew
            mrsGrIDData("行号").Value = intRow
            mrsGrIDData("列号").Value = intCol
            mrsGrIDData("内容").Value = vData
            mrsGrIDData("字体").Value = mrsGrIDFormat("字体").Value
            mrsGrIDData("大小").Value = mrsGrIDFormat("大小").Value
            mrsGrIDData("前景色").Value = mrsGrIDFormat("前景色").Value
            mrsGrIDData("背景色").Value = mrsGrIDFormat("背景色").Value
            mrsGrIDData("水平对齐").Value = mrsGrIDFormat("水平对齐").Value
            mrsGrIDData("垂直对齐").Value = mrsGrIDFormat("垂直对齐").Value
            mrsGrIDData("自动换行").Value = 0
            mrsGrIDData("多行输出").Value = 0
            mrsGrIDData("自动适应").Value = 0
            mrsGrIDData("旋转角度").Value = 0
        End Select
    End If
    

        
End Property

Public Property Let RowCell(ByVal Property As CellProperty, ByVal intRow1 As Integer, ByVal intRow2 As Integer, ByVal vData As Variant)

    mrsGrIDData.Filter = ""
    mrsGrIDData.Filter = "行号>=" & intRow1 & " And 行号<=" & intRow2
    
    If mrsGrIDData.RecordCount > 0 Then
        mrsGrIDData.MoveFirst
        Do While Not mrsGrIDData.EOF
        
            Select Case Property
            Case CellProperty.AligmentH
                mrsGrIDData("水平对齐").Value = Val(vData)
            Case CellProperty.AligmentV
                mrsGrIDData("垂直对齐").Value = Val(vData)
            Case CellProperty.FontName
                mrsGrIDData("字体").Value = Trim(vData)
            Case CellProperty.FontSize
                mrsGrIDData("大小").Value = Val(vData)
            Case CellProperty.ForeColor
                mrsGrIDData("前景色").Value = Val(vData)
            End Select
            
            mrsGrIDData.MoveNext
        Loop
    End If
    
End Property

Public Property Let ColCell(ByVal Property As CellProperty, ByVal intCol1 As Integer, ByVal intCol2 As Integer, ByVal vData As Variant)

    mrsGrIDData.Filter = ""
    mrsGrIDData.Filter = "列号>=" & intCol1 & " And 列号<=" & intCol2
    
    If mrsGrIDData.RecordCount > 0 Then
        mrsGrIDData.MoveFirst
        Do While Not mrsGrIDData.EOF
        
            Select Case Property
            Case CellProperty.AligmentH
                mrsGrIDData("水平对齐").Value = Val(vData)
            Case CellProperty.AligmentV
                mrsGrIDData("垂直对齐").Value = Val(vData)
            Case CellProperty.FontName
                mrsGrIDData("字体").Value = Trim(vData)
            Case CellProperty.FontSize
                mrsGrIDData("大小").Value = Val(vData)
            Case CellProperty.ForeColor
                mrsGrIDData("前景色").Value = Val(vData)
            End Select
            
            mrsGrIDData.MoveNext
        Loop
    End If
    
End Property


Public Function SetGrIDRowFormat(ByVal strRowHeight As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strRowHeight, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGrIDRowFormat.AddNew
        mrsGrIDRowFormat("行号").Value = intLoop + 1
        mrsGrIDRowFormat("行高").Value = mclsLisReportLib.GetTwipsY(Val(varTmp(intLoop)))
        mrsGrIDRowFormat("线宽").Value = 1
        mrsGrIDRowFormat("线型").Value = 0
    Next
    
    SetGrIDRowFormat = True
End Function

Public Function SetGrIDRowHeight(ByVal intRow As Integer, ByVal sglHeight As Single) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mrsGrIDRowFormat.Filter = ""
    mrsGrIDRowFormat.Filter = "行号=" & intRow + 1
    If mrsGrIDRowFormat.RecordCount > 0 Then
        mrsGrIDRowFormat("行高").Value = mclsLisReportLib.GetTwipsY(sglHeight)
    End If
    mrsGrIDRowFormat.Filter = ""

    SetGrIDRowHeight = True
    
End Function

Public Function SetGrIDRowLineType(ByVal intRow As Integer, ByVal bytLineType As Byte) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mrsGrIDRowFormat.Filter = ""
    mrsGrIDRowFormat.Filter = "行号=" & intRow
    If mrsGrIDRowFormat.RecordCount > 0 Then
        mrsGrIDRowFormat("线宽").Value = bytLineType
    End If
    
    SetGrIDRowLineType = True
End Function

Public Function SetGrIDRowLineStyle(ByVal intRow As Integer, ByVal bytLineStyle As Byte) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mrsGrIDRowFormat.Filter = ""
    mrsGrIDRowFormat.Filter = "行号=" & intRow
    If mrsGrIDRowFormat.RecordCount > 0 Then
        mrsGrIDRowFormat("线型").Value = bytLineStyle
    End If
    
    SetGrIDRowLineStyle = True
End Function

Public Function SetGrIDColFormat(ByVal strColWidth As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strColWidth, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGrIDColFormat.AddNew
        mrsGrIDColFormat("列号").Value = intLoop + 1
        mrsGrIDColFormat("列宽").Value = mclsLisReportLib.GetTwipsY(Val(varTmp(intLoop)))
        mrsGrIDColFormat("线宽").Value = 1
        mrsGrIDRowFormat("线型").Value = 0
    Next
    
    SetGrIDColFormat = True
    
End Function

Public Function SetGrIDColLineType(ByVal intCol As Integer, ByVal bytLineType As Byte) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mrsGrIDColFormat.Filter = ""
    mrsGrIDColFormat.Filter = "列号=" & intCol
    If mrsGrIDColFormat.RecordCount > 0 Then
        mrsGrIDColFormat("线宽").Value = bytLineType
    End If
    
    SetGrIDColLineType = True
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
    mrsGrIDColFormat.Filter = ""
    intCols = mrsGrIDColFormat.RecordCount
    For intLoop = 1 To intCols

        mrsGrIDColFormat.Filter = ""
        mrsGrIDColFormat.Filter = "列号=" & intLoop
        If mrsGrIDColFormat.RecordCount > 0 Then
'
            sglWidth = sglWidth + Val(mrsGrIDColFormat("列宽").Value)

        End If
    Next
    
    Width = mclsLisReportLib.GetCentimetersX(sglWidth)
End Function

Public Function SumGrIDColWidth(Optional ByVal intCol As Integer) As Single
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglWidth As Single
    Dim intCols As Integer
    
    sglWidth = 0
    mrsGrIDColFormat.Filter = ""
    intCols = mrsGrIDColFormat.RecordCount
    For intLoop = 1 To intCols

        mrsGrIDColFormat.Filter = ""
        mrsGrIDColFormat.Filter = "列号=" & intLoop
        If mrsGrIDColFormat.RecordCount > 0 Then
'
            sglWidth = sglWidth + Val(mrsGrIDColFormat("列宽").Value)
            
            If intLoop = intCol Then
                SumGrIDColWidth = sglWidth
                Exit Function
            End If

        End If
    Next
    
    SumGrIDColWidth = sglWidth
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
    
    mrsGrIDRowFormat.Filter = ""
    intRows = mrsGrIDRowFormat.RecordCount
    For intLoop = 1 To intRows

        mrsGrIDRowFormat.Filter = ""
        mrsGrIDRowFormat.Filter = "行号=" & intLoop
        If mrsGrIDRowFormat.RecordCount > 0 Then
            
            sglHeight = sglHeight + mclsLisReportLib.GetCentimetersY(Val(mrsGrIDRowFormat("行高").Value))
            
            If intLoop = intRow Then
                Height = sglHeight
                Exit Function
            End If
            
        End If
    Next
    
    Height = sglHeight
End Function

Public Function SumGrIDRowHeight(Optional ByVal intRow As Integer) As Single
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglHeight As Single
    Dim intRows As Integer
    
    sglHeight = 0
    
    mrsGrIDRowFormat.Filter = ""
    intRows = mrsGrIDRowFormat.RecordCount
    For intLoop = 1 To intRows

        mrsGrIDRowFormat.Filter = ""
        mrsGrIDRowFormat.Filter = "行号=" & intLoop
        If mrsGrIDRowFormat.RecordCount > 0 Then
'
            sglHeight = sglHeight + Val(mrsGrIDRowFormat("行高").Value)
            
            If intLoop = intRow Then
                SumGrIDRowHeight = sglHeight
                Exit Function
            End If

        End If
    Next
    
    SumGrIDRowHeight = sglHeight
End Function

Public Function SetGrIDColData(ByVal intCol As Integer, ByVal strData As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strData, ";")
    
    
    For intLoop = 0 To UBound(varTmp)
        mrsGrIDData.AddNew
        mrsGrIDData("行号").Value = intLoop + 1
        mrsGrIDData("列号").Value = intCol
        mrsGrIDData("内容").Value = CStr(varTmp(intLoop))
        mrsGrIDData("字体").Value = mrsGrIDFormat("字体").Value
        mrsGrIDData("大小").Value = mrsGrIDFormat("大小").Value
        mrsGrIDData("前景色").Value = mrsGrIDFormat("前景色").Value
        mrsGrIDData("背景色").Value = mrsGrIDFormat("背景色").Value
        mrsGrIDData("水平对齐").Value = mrsGrIDFormat("水平对齐").Value
        mrsGrIDData("垂直对齐").Value = mrsGrIDFormat("垂直对齐").Value
        mrsGrIDData("自动换行").Value = 0
        mrsGrIDData("多行输出").Value = 0
        mrsGrIDData("自动适应").Value = 0
        mrsGrIDData("旋转角度").Value = 0
'        mrsGrIDData("图形文件").Value = ""
    Next
    
    SetGrIDColData = True
    
End Function

Public Function SetGrIDRowData(ByVal intRow As Integer, ByVal strData As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strData, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGrIDData.AddNew
        mrsGrIDData("行号").Value = intRow
        mrsGrIDData("列号").Value = intLoop + 1
        mrsGrIDData("内容").Value = CStr(varTmp(intLoop))
        mrsGrIDData("字体").Value = mrsGrIDFormat("字体").Value
        mrsGrIDData("大小").Value = mrsGrIDFormat("大小").Value
        mrsGrIDData("前景色").Value = mrsGrIDFormat("前景色").Value
        mrsGrIDData("背景色").Value = mrsGrIDFormat("背景色").Value
        mrsGrIDData("水平对齐").Value = mrsGrIDFormat("水平对齐").Value
        mrsGrIDData("垂直对齐").Value = mrsGrIDFormat("垂直对齐").Value
        mrsGrIDData("自动换行").Value = 0
        mrsGrIDData("多行输出").Value = 0
        mrsGrIDData("自动适应").Value = 0
        mrsGrIDData("旋转角度").Value = 0
'        mrsGrIDData("图形文件").Value = ""
    Next
    
    SetGrIDRowData = True
    
End Function

Friend Function FormatData(Optional ByVal strKind As String = "数据") As Boolean
    '******************************************************************************************************************
    '功能： 产生表格数据
    '参数： mrsGrIDData　- 单元格数据；mrsGrIDRowFormat -  行格式说明；mrsGrIDColFormat - 列格式说明；
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
    Dim varGrIDLine As Variant
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim objSvrFont As USERFONT
    Dim objSvrRect As USERRECT
    
    On Error GoTo errHand
    
    Call SaveFont(gobjFont, objSvrFont)
    Call SaveRect(gobjRect, objSvrRect)
    
    X0 = gobjRect.X0
    Y0 = gobjRect.Y0
    X1 = gobjRect.X1
    Y1 = gobjRect.Y1
    
    mrsGrIDRowFormat.Filter = ""
    mrsGrIDColFormat.Filter = ""
    
    intRows = mrsGrIDRowFormat.RecordCount
    intCols = mrsGrIDColFormat.RecordCount
    
    '先画表格线
    '------------------------------------------------------------------------------------------------------------------
    varGrIDLine = Split(mrsGrIDFormat("表格线").Value, ";")
    
    '外框：上横线
    If Val(varGrIDLine(0)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.X1 = X1
        gobjRect.Y1 = Y0
        Call AppendPrintData(strKind, "线条")
    End If
    
    '外框：右竖线
    If Val(varGrIDLine(1)) > 0 Then
        gobjRect.X0 = X1
        gobjRect.Y0 = Y0
        gobjRect.X1 = X1
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "线条")
    
    End If
    
    '外框：下横线
    If Val(varGrIDLine(2)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y1
        gobjRect.X1 = X1
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "线条")
    
    End If
    
    '外框：左竖线
    If Val(varGrIDLine(3)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.X1 = X0
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "线条")
            
    End If
    
    
    '内框：横线
    gobjFont.LineWidth = 1
    If Val(varGrIDLine(4)) > 0 Then
        gobjRect.Y0 = Y0
        gobjRect.X0 = X0
        gobjRect.X1 = X1
        For intRow = 1 To intRows - 1
            mrsGrIDRowFormat.Filter = ""
            mrsGrIDRowFormat.Filter = "行号=" & intRow
            If mrsGrIDRowFormat.RecordCount > 0 Then
                gobjRect.Y0 = gobjRect.Y0 + Val(mrsGrIDRowFormat("行高").Value)
                gobjRect.Y1 = gobjRect.Y0
                If Val(mrsGrIDRowFormat("线宽").Value) <> 0 Then
                    gobjFont.LineWidth = Val(mrsGrIDRowFormat("线宽").Value)
                    gobjFont.LineStyle = Val(mrsGrIDRowFormat("线型").Value)
                    Call AppendPrintData(strKind, "线条")
                End If
            End If
        Next
    End If
        
    '内框：竖线
    If Val(varGrIDLine(5)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.Y1 = Y1
        For intCol = 1 To intCols - 1
            mrsGrIDColFormat.Filter = ""
            mrsGrIDColFormat.Filter = "列号=" & intCol
            If mrsGrIDColFormat.RecordCount > 0 Then
                gobjRect.X0 = gobjRect.X0 + Val(mrsGrIDColFormat("列宽").Value)
                gobjRect.X1 = gobjRect.X0
                gobjFont.LineWidth = Val(mrsGrIDColFormat("线宽").Value)
                gobjFont.LineStyle = Val(mrsGrIDRowFormat("线型").Value)
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
        
        mrsGrIDRowFormat.Filter = ""
        mrsGrIDRowFormat.Filter = "行号=" & intRow
        
        If mrsGrIDRowFormat.RecordCount > 0 Then

            gobjRect.Y1 = gobjRect.Y0 + Val(mrsGrIDRowFormat("行高").Value)
            gobjRect.X1 = X0
            
            For intCol = 1 To intCols
                
                mrsGrIDColFormat.Filter = ""
                mrsGrIDColFormat.Filter = "列号=" & intCol
                    
                If mrsGrIDColFormat.RecordCount > 0 Then
                
                    gobjRect.X0 = gobjRect.X1
                    gobjRect.X1 = gobjRect.X0 + Val(mrsGrIDColFormat("列宽").Value)
                                        
                    mrsGrIDData.Filter = ""
                    mrsGrIDData.Filter = "行号=" & intRow & " And 列号=" & intCol
                    If mrsGrIDData.RecordCount > 0 Then
     
                        strTmp = mrsGrIDData("内容").Value
                    
                        gobjFont.Name = mrsGrIDData("字体").Value
                        gobjFont.Size = mrsGrIDData("大小").Value
                        gobjFont.BackColor = Val(mrsGrIDData("背景色").Value)
                        gobjFont.ForeColor = Val(mrsGrIDData("前景色").Value)

                        gobjFont.Bold = False
                        gobjFont.Underline = False
                        Call SetDraw(gobjFont)
                        
                        '<Image>c:\a<Image>"
                        intPos1 = InStr(strTmp, "<Image>")
                        intPos2 = InStr(strTmp, "</Image>")
                        If intPos1 > 0 And intPos2 > 0 And intPos1 < intPos2 Then
                            strTmp = Mid(strTmp, intPos1 + 7, intPos2 - intPos1 - 7)
                            If Dir(strTmp) <> "" And strTmp <> "" Then
                                Call AppendPrintData(strKind, "图形", Val(mrsGrIDData("水平对齐").Value), False, strTmp)
                            End If
                        Else
                        
                            If Val(mrsGrIDData("水平对齐").Value) = 1 Then gobjRect.X0 = gobjRect.X0 + mclsLisReportLib.GetTwipsX(0.05)
                            If Val(mrsGrIDData("垂直对齐").Value) = 1 Then gobjRect.Y0 = gobjRect.Y0 + mclsLisReportLib.GetTwipsY(0.05)
                        
                            Call AppendPrintData(strKind, "文本", Val(mrsGrIDData("水平对齐").Value), (Val(mrsGrIDData("自动换行").Value) = 1), strTmp, Val(mrsGrIDData("垂直对齐").Value), (Val(mrsGrIDData("多行输出").Value) = 1), , (Val(mrsGrIDData("自动适应").Value) = 1), , , Val(mrsGrIDData("旋转角度").Value))
                            
                        
                            If Val(mrsGrIDData("水平对齐").Value) = 1 Then gobjRect.X0 = gobjRect.X0 - mclsLisReportLib.GetTwipsX(0.05)
                            If Val(mrsGrIDData("垂直对齐").Value) = 1 Then gobjRect.Y0 = gobjRect.Y0 - mclsLisReportLib.GetTwipsY(0.05)
                        
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
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call SaveFont(objSvrFont, gobjFont)
End Function

Private Sub Class_Initialize()
    Set mclsLisReportLib = New clsLisReportLib
End Sub

Private Sub Class_Terminate()
    Set mclsLisReportLib = Nothing
End Sub
