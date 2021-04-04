Attribute VB_Name = "mdlDocShiftBase"
Option Explicit
Public Const conMenu_DocShift_FilePopup = 1              '文件
Public Const conMenu_DocShift_PatiTypePopup = 2            '病人类型
Public Const conMenu_DocShift_PatiProjectPopup = 3              '病人项目
Public Const conMenu_DocShift_ViewPopup = 7             '查看
Public Const conMenu_DocShift_HelpPopup = 9            '帮助

'文件菜单
Public Const conMenu_DocShift_File_Preview = 101         '预览(&V)
Public Const conMenu_DocShift_File_Exit = 191            '退出(&X)

'病人类型菜单
Public Const conMenu_DocShift_Edit_New = 201         '新病人类型(&A)
Public Const conMenu_DocShift_Edit_Modify = 202          '修改(&M)
Public Const conMenu_DocShift_Edit_Delete = 203          '删除(&D)
Public Const conMenu_DocShift_Edit_Reuse = 204          '启用(&D)
Public Const conMenu_DocShift_Edit_Stop = 205          '停用(&D)

'病人项目菜单
Public Const conMenu_DocShift_Edit_NewProject = 301        '新病人项目(&A)
Public Const conMenu_DocShift_Edit_ModifyProject = 302          '修改(&M)
Public Const conMenu_DocShift_Edit_DeleteProject = 303          '删除(&D)
Public Const conMenu_DocShift_Edit_RowProject = 304        '序号相同的项目需合并

'查看菜单
Public Const conMenu_DocShift_View_ToolBar = 701              '工具栏(&T)
Public Const conMenu_DocShift_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const conMenu_DocShift_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const conMenu_DocShift_View_ToolBar_Size = 7013           '大图标(&B)
Public Const conMenu_DocShift_View_StatusBar = 702            '状态栏(&S)

'帮助菜单
Public Const conMenu_DocShift_Help_Help = 901        '帮助主题(&H)
Public Const conMenu_DocShift_Help_Web = 902         '&WEB上的中联
Public Const conMenu_DocShift_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_DocShift_Help_Web_Forum = 9023      '中联论坛(&F)
Public Const conMenu_DocShift_Help_Web_Mail = 9022       '发送反馈(&M)
Public Const conMenu_DocShift_Help_About = 991       '关于(&A)…

Public Function rsPatiType(ByVal strSName As String) As ADODB.Recordset
'根据病人简称获取病人类型信息
    
    On Error GoTo errH
    gstrSql = "Select 简称, 名称, 顺序,起始描述, 提取sql, 是否停用 From 医生交接班病人类型 Where 简称 = [1]"
    Set rsPatiType = zlDatabase.OpenSQLRecord(gstrSql, "获取病人类型信息", strSName)
    Exit Function
errH:
    MsgBox err.Description, vbInformation, "获取病人类型信息"
End Function

Public Function GetPatiTypeInfo(ByVal strType As String, Optional strPatiTypeInfo As String) As ADODB.Recordset
'根据病人类型简称获取病人类型信息
    
    gstrSql = ""
    If strPatiTypeInfo <> "" Then gstrSql = " And 项目名称=[2]"
    gstrSql = "Select 病人简称, 项目名称, 序号, 项目类别, Decode(输入形式, 1, '1-输入框', 2, '2-单项选择', 3, '3-多项选择') 输入形式," & vbNewLine & _
        "       Decode(Nvl(输入类型,0), 0, '0-文本', 1, '1-日期', 2, '2-数字') 输入类型, 输入格式, 输入值域, 输入行数," & vbNewLine & _
        "       Decode(提取来源, 1, '1-最新诊断', 2, '2-最新体征', 3, '3-输血情况', 4, '4-病历内容', '99', '99-SQL提取') 提取来源, 提取病历, 提取sql, 描述文字, 是否只读, 死亡则隐藏" & vbNewLine & _
        "From 医生交接班病人项目" & vbNewLine & _
        "Where 病人简称 = [1]" & gstrSql & vbNewLine & _
        "Order By 序号"
    Set GetPatiTypeInfo = zlDatabase.OpenSQLRecord(gstrSql, "获取病人类型信息", strType, strPatiTypeInfo)
End Function

Public Function GetSqlColor() As String
    Dim objfso As New FileSystemObject
    '公共方法:获取语法控件的SQL语法高亮显示设置
    '获取后直接将语法控件的SyntaxScheme属性设为返回值即可
    Dim strColor As String, strPath As String
    
    strPath = objfso.GetParentFolderName(GetSetting("ZLSOFT", "公共全局", "程序路径")) & "\PUBLIC\_sql.schclass"
    If Not objfso.FileExists(strPath) Then
        strPath = "C:\Appsoft\PUBLIC\_sql.schclass"
    End If
    
    If objfso.FileExists(strPath) Then
        strColor = ReadFileToString(strPath)
    End If
    GetSqlColor = strColor
End Function

Public Function ReadFileToString(ByVal strFile As String) As String
    Dim strBuffer As String
    Dim lngHwnd As Long
    Dim lngFileLen As Long

    lngHwnd = FreeFile

    On Error Resume Next
    Open strFile For Binary Shared As lngHwnd
    If err.Number <> 0 Then
        MsgBox "Error " & err.Number & vbCrLf & err.Description & vbCrLf & "Error in ReadFileToString, File='" & strFile & "'", vbCritical
        GoTo Proc_Exit
    End If
    On Error GoTo 0
    
    lngFileLen = LOF(lngHwnd)
    strBuffer = Space(lngFileLen)
    Get lngHwnd, , strBuffer
    
    Close lngHwnd
    
Proc_Exit:
    ReadFileToString = strBuffer
End Function


