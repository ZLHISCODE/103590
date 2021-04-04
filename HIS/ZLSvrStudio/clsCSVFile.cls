VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCSVFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'使用说明：
'1、读取CSV
'    Set objCsv = New clsCSVFile
'    If objCsv.LoadFile("C:\Test.csv") Then
'        For i = 0 To objCsv.Rows - 1
'            For j = 0 To objCsv.Cols - 1
'                Debug.Print objCsv.Text(i, j)
'            Next
'        Next
'    End If
'    objCsv.Text(2, 2)=3 '将第二行第三列修改为3
'    objCsv.ReSaveFile("C:\Test1.csv")'将已经读取或者修改后的CSV文件重新保存
'2、保存CSV
'    Set objCsv = New clsCSVFile
'    If Not objCsv.SaveRec("C:\Test.csv", rsTmp) Then
'        Debug.Print "ERROR"
'    End If

'    Set objCsv = New clsCSVFile
'    If Not objCsv.SaveStart("C:\Test.csv") Then
'        Debug.Print objCsv.ErrorDescription
'    End If
'    Call objCsv.SaveRow("1", "2", "3")  '输出第一行记录，1,2,3
'    Call objCsv.SaveValues("3", "4")    '输出第二行记录的前两列，3,4
'    Call objCsv.SaveRow                 '标记第二行输出完成，第二行第三列用空值填充，第二行内容为3,4,
'    Call objCsv.SaveRow("5", "6", "7", "8") '输出第三行记录，由于列数超过第一行列数，自动截断，只输出三列5,6,7
'    If Not objCsv.SaveEnd() Then
'        Debug.Print objCsv.ErrorDescription
'    End If
'3、同一个对象交错处理保存和加载会导致数据错误或程序错误。
'    If Not objCsv.SaveStart("C:\Test.csv") Then
'        Debug.Print objCsv.ErrorDescription
'    End If
'    If objCsv.LoadFile("C:\Test.csv") Then
'        For i = 0 To objCsv.Rows - 1
'            For j = 0 To objCsv.Cols - 1
'                Debug.Print objCsv.Text(i, j)
'            Next
'        Next
'    End If
'    Call objCsv.SaveRow("1", "2", "3")  '输出第一行记录，1,2,3
'若要使用一个对象操作保存和读取，请保证读取完成后保存、或保存完成后再读取，否则建议使用两个独立的对象
'属性值
Private mstrError       As String       'CSV加载或者保存的错误信息
Private mlngRows        As Long         'CSV加载或者保存时的行数
Private mlngCols        As Long         'CSV加载或者保存时的列数

Private mlngCurCol      As Long         '当前记录已经输出的列
Private mobjFSO         As New FileSystemObject
'SaveStart、SaveValues、SaveRow、SaveEnd使用的模块变量
Private mblnFirstRow    As Boolean      '是否是第一行
Private mobjSaveFile    As TextStream   '写文件的文件流

Private mclsBuffer      As New clsStringBulider
Private mobjTextLex     As TextStream
Private marrValues()    As String
Private mstrContent     As String
Private Enum StateType  '定义状态
    NewFieldStart = 0   '新字段开始
    NonQuotesField      '非引号字段
    QuotesField         '引号字段
    FieldSeparator      '字段分隔
    QuoteInQuotesField  '引号字段中的引号
    RowSeparator        '行分隔符字符1，回车
    LexError            '语法错误
End Enum
'======================================================================================================================
'属性           ErrorDescription       CSV加载或者保存的错误信息
'类型           String
'访问控制       (R)
'======================================================================================================================
Public Property Get ErrorDescription() As String
    ErrorDescription = mstrError
End Property
'======================================================================================================================
'属性           Rows                    CSV加载或者保存时的行数
'类型           Long
'访问控制       (R)
'======================================================================================================================
Public Property Get Rows() As Long
    Rows = mlngRows
End Property
'======================================================================================================================
'属性           Cols                    CSV加载或者保存时的列数
'类型           Long
'访问控制       (R)
'======================================================================================================================
Public Property Get Cols() As Long
    Cols = mlngCols
End Property
'======================================================================================================================
'属性           Text                    获取CSV某一行某一列的值
'类型           String
'访问控制       (R/W)
'======================================================================================================================
Public Property Get Text(Row As Long, Col As Long) As String
    If Row < mlngRows And Col < mlngCols And mlngCols > 0 Then
        Text = marrValues(Col + Row * mlngCols)
    End If
End Property

Public Property Let Text(Row As Long, Col As Long, ByVal strValues As String)
    If Row < mlngRows And Col < mlngCols And mlngCols > 0 Then
        marrValues(Col + Row * mlngCols) = strValues
    End If
End Property
'======================================================================================================================
'方法           LoadFile                加载读取CSV文件
'返回值         Boolean                 是否加载成功
'入参列表:
'参数名         类型                    说明
'strFileName    String                  CSV文件路径
'blnFile        Boolean                 传入的是否是文件名，否则是待解析的字符串
'======================================================================================================================
Public Function LoadFile(ByVal strFileName As String, Optional ByVal blnFile As Boolean = True) As Boolean
    On Error GoTo ErrH
    If Not mobjSaveFile Is Nothing Then
        Call SaveEnd
    End If
    ReDim Preserve marrValues(16)
    mlngRows = 0
    mlngCols = 0
    mblnFirstRow = True
    If blnFile Then
        Set mobjTextLex = mobjFSO.OpenTextFile(strFileName, ForReading)
    Else
        Set mobjTextLex = Nothing
    End If
    mstrContent = strFileName
    LoadFile = LexCSVContens
    If blnFile Then mobjTextLex.Close
    Set mobjTextLex = Nothing
    Exit Function
ErrH:
    mstrError = err.Description
End Function
'======================================================================================================================
'方法           ReSaveFile              重新保存从CSV文件中读取的文件
'返回值         Boolean                 是否保存成功
'入参列表:
'参数名         类型                    说明
'strFileName    String                  CSV文件路径
'======================================================================================================================
Public Function ReSaveFile(ByVal strFileName As String) As Boolean
    Dim i As Long, j As Long
    If Not mobjSaveFile Is Nothing Then
        Call SaveEnd
    End If
    If Not SaveStart(strFileName) Then
        Exit Function
    End If
    For i = 0 To Me.Rows - 1
        For j = 0 To Me.Cols - 1
            If Not SaveOneValue(Me.Text(i, j)) Then
                Exit Function
            End If
        Next
        If Not SaveRow() Then
            Exit Function
        End If
    Next
    If Not SaveEnd() Then
        Exit Function
    End If
    ReSaveFile = True
End Function
'======================================================================================================================
'方法           SaveRec                 将记录集保存为CSV文件
'返回值         Boolean                 是否保存成功
'入参列表:
'参数名         类型                    说明
'strFileName    String                  CSV文件路径
'rsSource       ADODB.Recordset         要保存的记录集
'说明：记录集从当前位置开始保存，若要保存全部记录，请清空Filter或者MoveFirst
'======================================================================================================================
Public Function SaveRec(ByVal strFileName As String, ByVal rsSource As ADODB.Recordset) As Boolean
    Dim i       As Long
    
    If Not SaveStart(strFileName) Then
        Exit Function
    End If
    If Not rsSource Is Nothing Then
        Do While Not rsSource.EOF
            For i = 0 To rsSource.Fields.Count - 1
                If Not SaveOneValue(rsSource.Fields(i).value & "") Then
                    Exit Function
                End If
            Next
            If Not SaveRow() Then
                Exit Function
            End If
            rsSource.MoveNext
        Loop
    End If
    If Not SaveEnd() Then
        Exit Function
    End If
    SaveRec = True
End Function

'======================================================================================================================
'方法           SaveStart            单行保存CSV文件的起始函数，SaveStart、SaveRow、SaveEnd请配套使用
'返回值         Boolean                 是否创建CSV文件成功
'入参列表:
'参数名         类型                    说明
'strFileName    String                  CSV文件路径
'======================================================================================================================
Public Function SaveStart(ByVal strFileName As String, Optional ByVal blnOwverWrite As Boolean = True) As Boolean
    On Error GoTo ErrH
     If Not mobjSaveFile Is Nothing Then
        Call SaveEnd
    End If
    mlngRows = 0
    mlngCols = 0
    mblnFirstRow = True
    Set mobjSaveFile = mobjFSO.OpenTextFile(strFileName, IIf(blnOwverWrite, ForWriting, ForAppending), True)
    SaveStart = True
    Exit Function
ErrH:
    mstrError = err.Description
    err.Clear
End Function
'======================================================================================================================
'方法           CSVSaveValue            保存单个或多个值到CSV文件
'返回值         Boolean                 是否保存成功
'入参列表:
'参数名         类型                    说明
'arrInput()     Variant                 要保存的值，或者保存的多个值的数组
'说明：该函数和SaveRow的区别在于，SaveRow认为完整输出的是一行记录。CSVSaveValue认为输出的是一行记录中的部分值，需要配合调用SaveRow来完成输出的行记录
'======================================================================================================================
Public Function SaveValues(ParamArray arrInput() As Variant) As Boolean
    Dim i       As Long, j  As Long

    For i = LBound(arrInput) To UBound(arrInput)
        If (varType(arrInput(i)) And vbArray) = vbArray Then
            For j = LBound(arrInput(i)) To UBound(arrInput(i))
                If Not SaveOneValue(CStr(arrInput(i)(j))) Then
                    Exit Function
                End If
            Next
        Else
            If Not SaveOneValue(CStr(arrInput(i))) Then
                Exit Function
            End If
        End If
    Next
    SaveValues = True
End Function

'======================================================================================================================
'方法           SaveOneValue         保存单个值到CSV文件
'返回值         Boolean                 是否保存成功
'入参列表:
'参数名         类型                    说明
'strInput       String                  要保存的值，或者保存的多个值的数组
'======================================================================================================================
Private Function SaveOneValue(strInput As String) As Boolean
    
    If mblnFirstRow Then
        mlngCols = mlngCols + 1
    ElseIf mlngCurCol >= mlngCols Then
        SaveOneValue = True
        mstrError = "(警告)行：" & mlngRows & "输出列数超过：" & mlngCols
        Exit Function
    End If
    mlngCurCol = mlngCurCol + 1
    If mlngCurCol > 1 Then
        mobjSaveFile.Write "," & ValueAdjust(strInput)
    Else
        mobjSaveFile.Write ValueAdjust(strInput)
    End If
    SaveOneValue = True
End Function
'======================================================================================================================
'方法           SaveRow              保存单行到CSV文件
'返回值         Boolean                 是否保存成功
'入参列表:
'参数名         类型                    说明
'arrInput()     Variant                  要保存的单行的每列，若只有一个参数，参数类型为数组类型，则将该数组的元素视为每列，数组元素尽量为字符串类型
'说明：CSV的每行列数必须相同，以第一次调用该函数的列数作为缺省列数，不足补齐，多了截断。
'      当没有传递任何参数，则默认输出一个换行符
'======================================================================================================================
Public Function SaveRow(ParamArray arrInput() As Variant) As Boolean
    Dim i       As Long, j  As Long

    For i = LBound(arrInput) To UBound(arrInput)
        If (varType(arrInput(i)) And vbArray) = vbArray Then
            For j = LBound(arrInput(i)) To UBound(arrInput(i))
                If Not SaveOneValue(CStr(arrInput(i)(j))) Then
                    Exit Function
                End If
            Next
        Else
            If Not SaveOneValue(CStr(arrInput(i))) Then
                Exit Function
            End If
        End If
    Next
    If mblnFirstRow Then
        mblnFirstRow = False
    ElseIf mlngCurCol < mlngCols Then
        For i = mlngCurCol + 1 To mlngCols
            If Not SaveOneValue("") Then
                Exit Function
            End If
        Next
        mstrError = "(警告)行：" & mlngRows & "输出列数不足：" & mlngCols
    End If
    mobjSaveFile.WriteLine
    mlngRows = mlngRows + 1
    mlngCurCol = 0
    SaveRow = True
End Function
'======================================================================================================================
'方法           SaveEnd              完成CSV文件的输出
'返回值         Boolean                 是否保存成功
'入参列表:
'参数名         类型                    说明
'======================================================================================================================
Public Function SaveEnd() As Boolean
    mobjSaveFile.Close
    Set mobjSaveFile = Nothing
    SaveEnd = True
End Function

'======================================================================================================================
'方法           ValueAdjust             调整为保存的CSV格式
'返回值         String                  CSV格式字符
'入参列表:
'参数名         类型                    说明
'strInput       String                  需要转换为CSV的值
'======================================================================================================================
Private Function ValueAdjust(ByRef strInput As String) As String
    If InStr(strInput, """") > 0 Then
        ValueAdjust = """" & Replace(strInput, """", """""") & """"
    Else
        If InStr(strInput, ",") > 0 Then
            ValueAdjust = """" & strInput & """"
        ElseIf InStr(strInput, vbCr) > 0 Then
            ValueAdjust = """" & strInput & """"
        ElseIf InStr(strInput, vbLf) > 0 Then
            ValueAdjust = """" & strInput & """"
        Else
            ValueAdjust = strInput
        End If
    End If
End Function

'======================================================================================================================
'方法           LexCSVContens           CSV语法解析器,使用有限状态机
'返回值         String                  语法解析后输出的结果
'入参列表:
'参数名         类型                    说明
'======================================================================================================================
Private Function LexCSVContens() As Boolean
    Dim strLine     As String
    Dim i           As Long
    Dim strOneChar  As String
    Dim stCur       As StateType
    
    strLine = ReadLine & vbCrLf
    mclsBuffer.Clear
    i = 1
    Do
        strOneChar = Mid(strLine, i, 1)
        Select Case stCur
            Case NewFieldStart
                Select Case strOneChar
                    Case """"
                        stCur = QuotesField
                    Case ","
                        Call ReadOneValue("")
                        stCur = FieldSeparator
                    Case vbCr, vbLf
                        mstrError = "存在空行"
                        stCur = LexError
                    Case Else
                        mclsBuffer.Append strOneChar
                        stCur = NonQuotesField
                End Select
            Case NonQuotesField
                Select Case strOneChar
                    Case ","
                        Call ReadOneValue(mclsBuffer.ToString())
                        mclsBuffer.Clear
                        stCur = FieldSeparator
                    Case vbCr
                        Call ReadOneValue(mclsBuffer.ToString())
                        stCur = RowSeparator
                    Case Else
                        mclsBuffer.Append strOneChar
                End Select
            Case QuotesField
                Select Case strOneChar
                    Case """"
                        stCur = QuoteInQuotesField
                    Case Else
                        mclsBuffer.Append strOneChar
                End Select
            Case FieldSeparator
                Select Case strOneChar
                    Case ","
                        Call ReadOneValue("")
                    Case """"
                        mclsBuffer.Clear
                        stCur = QuotesField
                    Case vbCr
                        Call ReadOneValue("")
                        stCur = RowSeparator
                    Case Else
                        mclsBuffer.Append strOneChar
                        stCur = NonQuotesField
                End Select
            Case QuoteInQuotesField
                Select Case strOneChar
                    Case ","
                        Call ReadOneValue(mclsBuffer.ToString())
                        mclsBuffer.Clear
                        stCur = FieldSeparator
                    Case vbCr
                        Call ReadOneValue(mclsBuffer.ToString())
                        stCur = RowSeparator
                    Case """" '转义字符
                        mclsBuffer.Append strOneChar
                        stCur = QuotesField
                    Case Else
                        mstrError = "语法错误： 转义字符""不能完成转义或引号字段结尾引号没有紧贴字段分隔符"
                        stCur = LexError
                End Select
            Case RowSeparator
                Select Case strOneChar
                    Case vbLf
                        '读取完成一行
                        mclsBuffer.Clear
                        stCur = NewFieldStart
                        If mblnFirstRow Then mblnFirstRow = False
                        mlngRows = mlngRows + 1
                        mlngCurCol = 0
                    Case Else
                        mstrError = "语法错误：行分隔用了回车vbCr。但未使用回车换行vbCrLf"
                        stCur = LexError
                End Select
        End Select
        If stCur = LexError Then
            Exit Do
        End If
        If i = Len(strLine) Then
            If stCur <> NewFieldStart Then
                If AtEndOfStream Then
                    mstrError = "语法错误：最后一行记录字段中包含回车vbCr或换行vbLf,但是未使用双引号括起来"
                    stCur = LexError
                    Exit Do
                End If
                strLine = strLine & ReadLine & vbCrLf
            Else
                i = 0
                If Not AtEndOfStream Then
                    strLine = ReadLine & vbCrLf
                Else '解析完成整个文件
                    Exit Do
                End If
            End If
        End If
        i = i + 1
    Loop While (i <= Len(strLine))
    LexCSVContens = stCur <> LexError
    Exit Function
End Function
'======================================================================================================================
'方法           ReadOneValue            读取一个字符并缓存
'返回值         Boolean                 是否保存成功
'入参列表:
'参数名         类型                    说明
'strInput       String                  要保存的值
'======================================================================================================================
Private Sub ReadOneValue(ByRef strValue As String)
    If mlngCurCol + mlngRows * mlngCols > UBound(marrValues) Then
        ReDim Preserve marrValues(UBound(marrValues) + 16 * (mlngCols + 1))
    End If
    If mblnFirstRow Then
        marrValues(mlngCurCol) = strValue
        mlngCols = mlngCols + 1
        mlngCurCol = mlngCurCol + 1
    ElseIf mlngCurCol < mlngCols Then
        marrValues(mlngCurCol + mlngRows * mlngCols) = strValue
        mlngCurCol = mlngCurCol + 1
    Else
        mstrError = "(警告)行：" & mlngRows & "读取列数超过：" & mlngCols
    End If
End Sub

Private Function ReadLine() As String
    If mobjTextLex Is Nothing Then
        ReadLine = mstrContent
    Else
        ReadLine = mobjTextLex.ReadLine
    End If
End Function

Private Function AtEndOfStream() As Boolean
    If mobjTextLex Is Nothing Then
        AtEndOfStream = True
    Else
        AtEndOfStream = mobjTextLex.AtEndOfStream
    End If
End Function

Private Sub Class_Initialize()
    ReDim Preserve marrValues(16)
End Sub

Private Sub Class_Terminate()
    If Not mobjSaveFile Is Nothing Then
        Call SaveEnd
    End If
    If Not mobjTextLex Is Nothing Then
        mobjTextLex.Close
        Set mobjTextLex = Nothing
    End If
    Set mclsBuffer = Nothing
    Set mobjFSO = Nothing
    Erase marrValues
End Sub


