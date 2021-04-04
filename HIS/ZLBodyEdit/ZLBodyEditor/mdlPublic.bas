Attribute VB_Name = "mdlPublic"
Option Explicit

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

Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'################################################################################################################
'##  得到用户的信息
'################################################################################################################
Public Sub GetUserInfo()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    On Error GoTo errHand
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
errHand:
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
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
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
    
    gstrSQL = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]" & IIf(blnMoved, ",1", "") & ") as 片段 From Dual"
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

errHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
    If ErrCenter = 1 Then
        Resume
    End If
End Function


'################################################################################################################
'## 功能：  将指定的文件保存到指定记录的LOB字段中
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  成功返回True，失败返回False
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobSave(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        gstrSQL = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "zlBlobSave")
    Next
    Close lngFileNum
    zlBlobSave = True
    Exit Function

errHand:
    Close lngFileNum
    zlBlobSave = False
End Function
Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub UniteCellCol(ByVal objCell As Object, _
                        ByVal intCOl As Integer, _
                        ByVal intRow As Integer, _
                        Optional startCol As Integer = 1)

    '功能：合并单元格的列.
    '参数: intcol 要合并的列数  introw第几行  startCol 起始列
    On Error GoTo errHand

    Dim strText As String
    Dim i As Integer, j As Integer

    With objCell
        .MergeRow(intRow) = True
        strText = Space(intRow + startCol) & Space(intRow + startCol)
        
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
                strText = Space(intRow + i) & Space(intRow + i)
            End If
        Next i
    End With

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub


