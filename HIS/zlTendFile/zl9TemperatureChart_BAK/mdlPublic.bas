Attribute VB_Name = "mdlPublic"
Option Explicit
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Public Const tomAutoColor As Long = -9999997
Public gstrFields As String
Public gstrValues As String


Public gstrProductName As String            '产品简称，例如：中联
Public gstrSysName As String                '系统名称，例如：中联软件
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrMatch As String                  '根据本地参数“匹配模式”确定的左匹配符号

Public gstrDbOwner As String                '当前数据库所有者（不同模块可能不一样）
Public gstrDBUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public glngSys As Long
Public gstrSQL As String
Public gcnOracle As New ADODB.Connection
Public gobjTendEditor As Object

Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'################################################################################################################
'##  得到用户的信息
'################################################################################################################
Public Sub GetUserInfo()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    On Error GoTo Errhand
    strSQL = "select u.用户名, P.*,D.编码 as 部门编码,D.名称 as 部门名称,M.部门ID" & _
                " from 上机人员表 U,人员表 P,部门表 D,部门人员 M " & _
                " Where U.人员id = P.id And P.ID=M.人员ID and  M.缺省=1 and M.部门id = D.id and U.用户名=user And (p.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.撤档时间 Is Null) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetUserInfo")
    With rsTemp
        If .RecordCount <> 0 Then
            gstrDBUser = .Fields("用户名").Value
            glngUserId = .Fields("ID").Value                '当前用户id
            gstrUserCode = .Fields("编号").Value            '当前用户编码
            gstrUserName = .Fields("姓名").Value            '当前用户姓名
            gstrUserAbbr = IIf(IsNull(.Fields("简码").Value), "", .Fields("简码").Value)          '当前用户简码
            glngDeptId = .Fields("部门id").Value            '当前用户部门id
            gstrDeptCode = .Fields("部门编码").Value        '当前用户
            gstrDeptName = .Fields("部门名称").Value        '当前用户
        Else
            gstrDBUser = ""
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
        .Close
    End With
   
   
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Sub

'################################################################################################################
'## 功能：  将指定的LOB字段复制为临时文件
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  存放内容的文件名，失败则返回零长度""
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, _
                           ByVal KeyWord As String, _
                           Optional ByRef strFile As String, _
                           Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240

    Dim lngFileNum     As Long, lngCount As Long, lngBound As Long

    Dim aryChunk()     As Byte, strText As String

    Dim rsLob          As New ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand
    
    lngFileNum = FreeFile

    If strFile = "" Then
        lngCount = 0

        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"

            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop

    End If

    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]) as 片段 From Dual"
    lngCount = 0

    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", lngCount)

        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte

        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop

    Close lngFileNum

    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile

    Exit Function

Errhand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UniteCellCol(ByVal objCell As Object, _
                        ByVal intCOl As Integer, _
                        ByVal intRow As Integer, _
                        Optional startCol As Integer = 1)

    '功能：合并单元格的列.
    '参数: intcol 要合并的列数  introw第几行  startCol 起始列
    On Error GoTo Errhand

    Dim strText As String
    Dim i As Integer, j As Integer

    With objCell
        .MergeRow(intRow) = True
        strText = " " & String(intRow, " ")
        
        For i = startCol To .Cols - 2
        
            j = i - objCell.FixedCols + 1
            If j < 0 Then j = 1
            
            If j Mod intCOl <> 0 Then
                .MergeCol(i) = True
                .Row = intRow
                .Col = i
                .Text = strText
                .CellAlignment = 4
                .Row = intRow
                .Col = i + 1
                .Text = strText
                .CellAlignment = 4
            Else
                strText = String(j / intCOl + 1, " ") & String(intRow, " ")
            End If
        Next i
    End With

    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------------------
'以下是基础函数或过程
'---------------------------------------------------------------------------------
Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, _
                      ByVal strFields As String, _
                      ByVal strValues As String)

    Dim arrFields, arrValues, intField As Integer

    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)

    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew

        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next

        .Update
    End With

End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, _
                         ByVal strFields As String, _
                         ByVal strValues As String, _
                         ByVal strPrimary As String, _
                         Optional ByVal blnDelete As Boolean = False)

    Dim arrFields, arrValues, intField As Integer

    '更新记录,如果不存在,则新增
    'strPrimary:字段名,值
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)

    If intField < 0 Then Exit Sub

    With rsObj

        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew

        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next

        .Update
    End With

End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, _
                              ByVal strPrimary As String, _
                              Optional ByVal blnDelete As Boolean = False) As Boolean

    Dim arrTmp

    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")

    With rsObj

        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"

        If .EOF Then Exit Function
        If blnDelete Then

            Do While Not .EOF

                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop

        Else
            Record_Locate = True
        End If

    End With

End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)

    Dim arrFields, intField As Integer

    Dim strFieldName As String, intType As Integer, lngLength As Long

    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj

        If .State = 1 Then .Close

        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then

                Select Case intType

                    Case adDouble
                        lngLength = madDoubleDefault

                    Case adVarChar
                        lngLength = madLongVarCharDefault

                    Case adLongVarChar
                        lngLength = madLongVarCharDefault

                    Case Else
                        lngLength = madDbDateDefault
                End Select

            End If

            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

End Sub

Public Sub OutputRsData(ByVal rsObj As ADODB.Recordset, _
                         Optional ByVal blnMod_Add As Boolean = False)

    Dim strOutput As String

    Dim intCOl    As Integer, intCols As Integer

    With rsObj

        If .RecordCount <> 0 Then .MoveFirst

        Do While Not .EOF
            strOutput = ""
            intCols = .Fields.Count

            For intCOl = 1 To intCols

                If Not blnMod_Add Then
                    strOutput = strOutput & "," & .Fields(intCOl - 1).Name & ":" & .Fields(intCOl - 1).Value
                Else
                    strOutput = strOutput & "|" & .Fields(intCOl - 1).Value
                End If

            Next

            Debug.Print Mid(strOutput, 2)
            
            .MoveNext
        Loop

        If .RecordCount <> 0 Then .MoveFirst
    End With

End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function LPAD(ByVal strText As String, ByVal intCount As Integer, ByVal strPAD As String) As String
'功能：等同Oracle的LPAD函数
    If LenB(StrConv(strText, vbFromUnicode)) < intCount Then
        LPAD = String(intCount - LenB(StrConv(strText, vbFromUnicode)), strPAD) & strText
    Else
        LPAD = strText
    End If
End Function



