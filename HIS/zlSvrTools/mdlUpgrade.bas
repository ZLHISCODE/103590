Attribute VB_Name = "mdlUpgrade"
Option Explicit

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Const NO_ERROR = 0
Const CONNECT_UPDATE_PROFILE = &H1
Const RESOURCETYPE_DISK = &H1
Const RESOURCETYPE_PRINT = &H2
Const RESOURCETYPE_ANY = &H0
Const RESOURCE_CONNECTED = &H1
Const RESOURCE_REMEMBERED = &H3
Const RESOURCE_GLOBALNET = &H2
Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Const RESOURCEDISPLAYTYPE_SERVER = &H2
Const RESOURCEDISPLAYTYPE_SHARE = &H3
Const RESOURCEUSAGE_CONNECTABLE = &H1
Const RESOURCEUSAGE_CONTAINER = &H2

Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias _
        "WNetAddConnection2A" _
        (lpNetResource As NETRESOURCE, _
        ByVal lpPassword As String, _
        ByVal lpUserName As String, _
        ByVal dwFlags As Long) As Long

Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias _
        "WNetCancelConnection2A" _
        (ByVal lpName As String, _
        ByVal dwFlags As Long, _
        ByVal fForce As Long) As Long


Public Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'功能：将指定INI配置文件的内容读取到记录集中
'返回：Nothing或包含"项目,内容"的记录集,其中同一项目可能有多行内容
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As Scripting.TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "项目", adVarChar, 100
    rsTmp.Fields.Append "内容", adVarChar, 4000, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        If Left(Trim(strLine), 1) = "[" And InStr(strLine, "]") > InStr(strLine, "[") Then
            
            If strItem <> "" And strText = "" Then
                rsTmp.AddNew
                rsTmp!项目 = strItem
                rsTmp!内容 = Null
                rsTmp.Update
            End If
            
            strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
            strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))
            If strItem <> "" And strText <> "" Then
                rsTmp.AddNew
                rsTmp!项目 = strItem
                rsTmp!内容 = strText
                rsTmp.Update
            End If
        ElseIf Trim(strLine) <> "" And strItem <> "" Then
            strText = Trim(strLine)
            rsTmp.AddNew
            rsTmp!项目 = strItem
            rsTmp!内容 = strText
            rsTmp.Update
        End If
    Loop
    
    If strItem <> "" And strText = "" Then
        rsTmp.AddNew
        rsTmp!项目 = strItem
        rsTmp!内容 = Null
        rsTmp.Update
    End If
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function

Public Function CheckINIValid(rsINI As ADODB.Recordset, ByVal strItem As String) As Boolean
'功能：检查对应的配置文件格式是否正确
'参数：rsINI=存放配置文件内容的记录集，包含"项目,内容"字段
'      strItem=配置文件中必须要求有内容的项目串,如"项目1|项目2|..."
    Dim arrItem As Variant, i As Long
    
    arrItem = Split(strItem, "|")
    For i = 0 To UBound(arrItem)
        rsINI.Filter = "项目='" & arrItem(i) & "'"
        If rsINI.EOF Then Exit Function
        If IsNull(rsINI!内容) Then Exit Function
    Next
    CheckINIValid = True
End Function

Public Function VerCompare(ByVal strVer1 As String, ByVal strVer2 As String, Optional ByVal blnPrimary As Boolean) As Integer
'功能：比较两个字符串表示的版本号的大小
'参数：blnPrimary=是否只比较"主版本.次版本",不管附版本
'返回：1=strVer1>strVer1,-1=strVer1<strVer1,0=strVer1=strVer1
'说明：VB最大支持的版本号为9999.9999.9999
    Dim arrVer As Variant
    
    If strVer1 Like "*.*.*" And strVer2 Like "*.*.*" Then
        arrVer = Split(strVer1, ".")
        strVer1 = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & IIf(blnPrimary, "", "." & Format(arrVer(2), "0000"))
        
        arrVer = Split(strVer2, ".")
        strVer2 = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & IIf(blnPrimary, "", "." & Format(arrVer(2), "0000"))
    End If
    If strVer1 > strVer2 Then
        VerCompare = 1
    ElseIf strVer1 < strVer2 Then
        VerCompare = -1
    End If
End Function

Public Function VerFull(ByVal strVer As String) As String
'功能：返回VB最大支持的版本号形式:9999.9999.9999
    Dim arrVer As Variant
    
    arrVer = Split(strVer, ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000")
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'功能：取指定字符串按字节算的长度
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function ActualStr(ByVal strAsk As String, ByVal lngLen As Long) As String
'功能：取指定字符串左边指定字节长度的内容
    Dim strTemp As String, i As Long
    
    strTemp = StrConv(LeftB(StrConv(strAsk, vbFromUnicode), lngLen), vbUnicode)
    If InStr(strTemp, Chr(0)) > 0 Then
        strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
    End If
    ActualStr = strTemp
End Function

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'说明：主要是RunSQLFile的子函数
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    Do While InStr(strText, "  ") > 0
        strText = Replace(strText, "  ", " ")
    Loop
    TrimEx = strText
End Function

Public Function TrimComment(ByVal strSQL As String) As String
'功能：去掉写在单行strSQL语句后面的"--"注释
'说明：主要是RunSQLFile的子函数
    Dim blnStr As Boolean
    Dim i As Long, k As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                k = i: Exit For
            End If
        Next
        If k > 0 Then strSQL = RTrim(Left(strSQL, k - 1))
    End If
    TrimComment = strSQL
End Function

Public Function SplitSQL(ByVal strSQL As String) As String
'功能：取";"结尾前面的的SQL语句,可能";"号后有"--"注释。
'说明：主要是RunSQLFile的子函数
    Dim i As Long, k As Long
    
    '先去掉注释部份
    strSQL = TrimComment(strSQL)
    
    For i = Len(strSQL) To 1 Step -1
        If Mid(strSQL, i, 1) = ";" Then
            k = i: Exit For
        End If
    Next
    If k > 0 Then strSQL = Left(strSQL, k - 1)
    
    SplitSQL = strSQL
End Function

Public Function RemoveMark(ByVal strText As String) As String
'功能：去除一段文字中的前导"--"注释标记
    Dim arrText As Variant, strTemp As String, i As Long
    
    arrText = Split(strText, vbCrLf)
    
    strText = ""
    For i = 0 To UBound(arrText)
        strTemp = arrText(i)
        If Left(strTemp, 2) = "--" And Replace(strTemp, "-", "") <> "" Then
            strText = strText & vbCrLf & Mid(strTemp, 3)
        End If
    Next
    RemoveMark = Mid(strText, 3)
End Function

Public Function GetLogSQL(objSQL As clsSQLInfo) As String
'功能：获取简要SQL语句，用于填写日志
    Dim strSQL As String
    
    If objSQL.Block Then
        If objSQL.BlockName <> "" Then
            strSQL = Trim(Split(objSQL.SQL, vbCrLf)(0))
            If InStr(strSQL, "(") > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(strSQL, "(") - 1))
            End If
            If InStr(1, strSQL, " as", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " as", vbTextCompare) - 1))
            End If
            If InStr(1, strSQL, " is", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " is", vbTextCompare) - 1))
            End If
            If InStr(1, strSQL, " Return", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " Return", vbTextCompare) - 1))
            End If
        Else '匿名块
            strSQL = ActualStr(TrimEx(objSQL.SQL, True), 150)
        End If
    ElseIf UCase(LTrim(objSQL.SQL)) Like "CREATE * VIEW" Then
        '视图特殊处理
        strSQL = Split(objSQL.SQL, vbCrLf)(0)
        If InStr(1, strSQL, " as", vbTextCompare) > 0 Then '视图只能用as
            strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " as", vbTextCompare) - 1))
        End If
    Else
        If InStr(objSQL.SQL, vbCrLf) > 0 Then
            '多行SQL
            strSQL = ActualStr(TrimEx(objSQL.SQL, True), 150)
        Else
            strSQL = ActualStr(objSQL.SQL, 150)
        End If
    End If
    GetLogSQL = strSQL
End Function


Public Function CheckHavHistory(ByVal lngSys As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:检查是否需要创建历史空间
    '参数:lngSys-系统号
    '返回:需要创建,返true,否则False
    '--------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 1 from zltools.zlbakTables where 系统=" & lngSys & " and rownum<=1"
    OpenRecordset rsTemp, gstrSQL, "获取bak数据", , , gcnOracle
    If rsTemp.EOF Then
       '返回False,表示该系统没有历史数据空间,没有要处理历史数据空间
       Exit Function
    End If
    CheckHavHistory = True
End Function

Public Function GrantBakToUser(ByVal cnOracle As ADODB.Connection, ByVal strToOwner As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------------
    '功能:检查表是否存在
    '参数:strTableName-表名
    '     cnoracle-数据库连接名
    '     strOwNer-所有者
    '返回:存在该表返回true,否则False
    '-----------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Err = 0: On Error GoTo ErrHand:
    strSQL = "Select TABLE_NAME from user_all_tables" & _
            " Union All Select View_Name From User_Views"
    Call OpenRecordset(rsTemp, strSQL, "重新授权", , , cnOracle)
    With rsTemp
        Do While Not .EOF
            strSQL = "Grant ALL on " & Nvl(!Table_Name) & " to " & strToOwner & " With Grant Option"
            cnOracle.Execute strSQL
            .MoveNext
        Loop
    End With
    GrantBakToUser = True
    Exit Function
ErrHand:
    If MsgBox("在授权时出现如下错误,请检查!" & vbCrLf & " (" & Err.Number & ") " & Err.Description, vbRetryCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbRetry Then
        Resume
    End If
    GrantBakToUser = False
End Function


Public Function IsNetServer(ByVal strPath As String, ByVal strUser As String, ByVal strPassWord As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '--功能:检查服务器是否正常并连接
    '--参数:strPath -访问路径
    '       strUser-用户名
    '       strPassWord -访问密码
    '返回:连接顺畅,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/06
    '----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
      
    '刘兴洪:可能存在windows资源管理器已经有访问的了
    '
    If objFile.FolderExists(strPath) Then
        IsNetServer = True: Exit Function
    End If
    
    Dim NetR As NETRESOURCE
    With NetR
        .dwScope = RESOURCE_GLOBALNET
        .dwType = RESOURCETYPE_DISK
        .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
        .dwUsage = RESOURCEUSAGE_CONNECTABLE
        .lpLocalName = "" '映射的驱动器
        .lpRemoteName = strPath  '服务器路径
    End With
    
    Err = 0
    On Error GoTo ErrHand:
    If WNetAddConnection2(NetR, strPassWord, strUser, CONNECT_UPDATE_PROFILE) = NO_ERROR Then
       IsNetServer = True
    Else
       IsNetServer = False
    End If
    Exit Function
ErrHand:
       IsNetServer = False
End Function
Public Function CancelNetServer(ByVal strPath As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '功能:断开服务器连接
    '参数:
    '返回:断找成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------------------------
    Err = 0
    On Error Resume Next
    If WNetCancelConnection2(strPath, CONNECT_UPDATE_PROFILE, True) = 0 Then
        CancelNetServer = True
    Else
        CancelNetServer = False
    End If
    Err = 0
End Function


