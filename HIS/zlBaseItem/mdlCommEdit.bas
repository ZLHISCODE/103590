Attribute VB_Name = "mdlCommEdit"
Option Explicit
Public gcnOracle As New ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public glngModul As Long

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录

Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstr单位名称 As String
Public gstrSQL As String
Public glngSys As Long

Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Sub GetUserInfo()
    '功能:得到用户的信息
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo ErrHand
    
    Set rsTemp = zlDatabase.GetUserInfo
    
    With rsTemp
        If .RecordCount <> 0 Then
            glngUserId = .Fields("ID").Value                '当前用户id
            gstrUserCode = .Fields("编号").Value            '当前用户编码
            gstrUserName = .Fields("姓名").Value            '当前用户姓名
            gstrUserAbbr = IIF(IsNull(.Fields("简码").Value), "", .Fields("简码").Value)          '当前用户简码
            glngDeptId = .Fields("部门id").Value            '当前用户部门id
            gstrDeptCode = .Fields("部门码").Value        '当前用户
            gstrDeptName = .Fields("部门名").Value        '当前用户
        Else
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Sub

Public Function GetDownCodeLength(ByVal strID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '功能描述：读取指定表的本级编码的最大长度
    '输入参数：本级ID，表名
    '输出参数：成功返回 下级最大编码; 否者返回 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If strID = "" Then
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " start with 上级ID is null " & strWhere & " connect by prior id=上级id"
    Else
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " start with 上级ID=" & strID & strWhere & " connect by prior id=上级id"
    End If
'    Call SQLTest(App.ProductName, "本级最大长度", strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetDownCodeLength")
'    Call SQLTest
    
    If rsTemp.RecordCount = 0 Then
        GetDownCodeLength = 0
    Else
        GetDownCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetDownCodeLength = 0
End Function

Public Function GetLocalCodeLength(ByVal str上级ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '功能描述：读取指定表的本级编码的最大长度
    '输入参数：上级ID，表名
    '输出参数：成功返回 最大编码; 否者返回 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " where 上级ID is null" & strWhere
    Else
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " where 上级ID=" & str上级ID & strWhere
    End If
'    Call SQLTest(App.ProductName, "本级最大长度", strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetLocalCodeLength")
'    Call SQLTest
    
    If rsTemp.RecordCount = 0 Then
        GetLocalCodeLength = 0
    Else
        GetLocalCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetLocalCodeLength = 0
End Function

Public Function GetParentCode(ByVal str上级ID As String, ByVal strTableName As String) As String
    '功能描述：读取上级编码
    '输入参数：上级ID,表名
    '输出参数：成功返回 上级编码; 否者返回 空
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        GetParentCode = ""
        Exit Function
    Else
        strSQL = "select 编码 from " & strTableName & " where ID=" & str上级ID
    End If
'    Call SQLTest(App.ProductName, "上级编码", strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetParentCode")
'    Call SQLTest
    
    If rsTemp.RecordCount = 0 Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("编码").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetParentCode = ""
End Function

Public Function GetMaxLocalCode(ByVal str上级ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
    '功能描述：根据指定表的上级ID 读取本级的最大编码
    '输入参数：上级ID,表名
    '输出参数：成功返回 最大编码; 否者返回 空
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim intCode As Integer, strCode As String, strAllCode As String
    Dim intLength   As Integer
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        strSQL = "select nvl(max(to_number(编码)),0)+1 as MaxCode from " & strTableName & " where 上级ID is null" & strWhere
        
        '如果是部门表，则要排除"已删除部门"分类的ID
        If strTableName = "部门表" Then
            strSQL = strSQL & " And 编码 <> '-'"
        End If
    Else
        strSQL = "select nvl(max(to_number(编码)),0)+1 as MaxCode from " & strTableName & " where 上级ID=" & str上级ID & strWhere
    End If
    intCode = GetLocalCodeLength(str上级ID, strTableName, strWhere)
'    Call SQLTest(App.ProductName, "本级最大编码", strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetMaxLocalCode")
'    Call SQLTest
    
    If rsTemp.EOF Then
        GetMaxLocalCode = ""
        Exit Function
    End If
    intLength = intCode - Len(IIF(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
    strAllCode = String(IIF(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
    'strCode = Mid(strAllCode, Len(GetParentCode(str上级ID, strTableName)) + 1)
    'GetMaxLocalCode = String(intCode - Len(strAllCode), "0") & strCode
    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str上级ID, strTableName)) + 1)
    If GetMaxLocalCode = "" Then GetMaxLocalCode = "1"
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetMaxLocalCode = ""
End Function

Public Function Where撤档时间(Optional strAlias As String) As String
    If strAlias = "" Then
        Where撤档时间 = " (撤档时间=to_date('3000-01-01','yyyy-mm-dd') or 撤档时间 is null) "
    Else
        Where撤档时间 = " (" & strAlias & ".撤档时间=to_date('3000-01-01','yyyy-mm-dd') or " & strAlias & ".撤档时间 is null) "
    End If
End Function

Public Function TruncateDate(ByVal datFull As Date) As Date
'去掉日期中的时、分、秒
    TruncateDate = CDate(Format(datFull, "yyyy-MM-dd"))
End Function

Public Function GetTextFromList(lstTemp As ListBox) As String
'参数：lstTemp  准备获取数据的ListBox控件
    Dim lngCount As Long
    Dim lngPos As Long
    Dim strTemp As String
    
    For lngCount = 0 To lstTemp.ListCount - 1
        If lstTemp.Selected(lngCount) = True Then
            lngPos = InStr(lstTemp.List(lngCount), ".")
            If lngPos = 0 Then
                strTemp = strTemp & lstTemp.List(lngCount) & ","
            Else
                strTemp = strTemp & Mid(lstTemp.List(lngCount), 1, lngPos - 1) & ","
            End If
        End If
    Next
    If strTemp <> "" Then
        '去掉最后一个,符号
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
    End If
    GetTextFromList = "'" & strTemp & "'"
End Function

Public Sub SetListByText(lstTemp As ListBox, ByVal strText As String)
'参数：lstTemp  准备设置的ListBox控件
    Dim lngCount As Long, lngIndex As Long, lngPos As Long
    Dim strTemp As String, varTemp As Variant
    Dim blnMatch As Boolean
    
    varTemp = Split(strText, ",")
    For lngCount = 0 To lstTemp.ListCount - 1
        blnMatch = False
        '取出该的值
        lngPos = InStr(lstTemp.List(lngCount), ".")
        If lngPos = 0 Then
            strTemp = lstTemp.List(lngCount)
        Else
            strTemp = Mid(lstTemp.List(lngCount), 1, lngPos - 1)
        End If
        For lngIndex = LBound(varTemp) To UBound(varTemp)
            If strTemp = varTemp(lngIndex) Then
                '已经找到相同的
                blnMatch = True
                Exit For
            End If
        Next
        lstTemp.Selected(lngCount) = blnMatch
    Next
End Sub


Public Sub ResetSelect(lvw As ListView, ByVal strKey As String)
'功能：重新设置ListView的选中项
'参数：strKey   刷新前的选中项
    Dim lst As ListItem
    
    If lvw.ListItems.Count > 0 Then
        On Error Resume Next
        Set lst = lvw.ListItems(strKey)
        If Err <> 0 Then
            '没有选中，也许该行已经被删除
            Err.Clear
            Set lst = lvw.ListItems(1)
        End If
        
        '设置选中项
        lst.Selected = True
        lst.EnsureVisible
    End If
End Sub

Public Sub RemoveSelect(lvw As ListView)
'功能：删除当前选中项
    Dim lngIndex  As Long
    
    With lvw
        If .SelectedItem Is Nothing Then Exit Sub
        
        lngIndex = .SelectedItem.Index
        .ListItems.Remove lngIndex
        
        If .ListItems.Count > 0 Then
            '如果仍有列表，则进行下一个选择
            lngIndex = IIF(.ListItems.Count > lngIndex, lngIndex, .ListItems.Count)
            .ListItems(lngIndex).Selected = True
            .ListItems(lngIndex).EnsureVisible
        End If
    End With

End Sub

