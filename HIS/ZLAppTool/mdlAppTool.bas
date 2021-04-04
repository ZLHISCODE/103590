Attribute VB_Name = "mdlAppTool"
Option Explicit
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = LVM_FIRST + 30
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const SW_SHOWNORMAL = 1

Public Type ChooseColorType
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     rgbResult As Long
     lpCustColors As String
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type

Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColorType) As Long
Public Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function SetActiveWindow Lib "User32" (ByVal hWnd As Long) As Long

Public Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Public gcnOracle As New ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gclsAppTool As clsAppTool       '当前APPTool对象
Public gstrPrivs As String                   '当前用户具有的当前模块的功能

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
Public gstrMenuSys As String                '当前用户使用的菜单系统
Public glngSys As Long                      '当前系统

'以下是消息系统要用到的全局变量
Public gfrmMain As Object                   '导航台窗口，主要用于作消息编辑窗口的父窗口
Public gblnMessageShow As Boolean           '说明消息主窗口是否已经显示
Public gblnMessageGet  As Boolean           '说明导航台是否要求通知新邮件

Public Const glngLBound As Long = 99
Public Const glngUBound As Long = 240

Public Sub GetUserInfo()
'功能:得到用户的信息

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
    
    rsTemp.CursorLocation = adUseClient
    On Error GoTo errHand
    
    With rsTemp
        strSQL = "select P.id,P.编号,P.姓名,P.简码,D.编码 as 部门编码,D.名称 as 部门名称,M.部门ID" & _
                " from 上机人员表 U,人员表 P,部门表 D,部门人员 M " & _
                " Where U.人员id = P.id And P.ID=M.人员ID and  M.缺省=1 and M.部门id = D.id and (P.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.撤档时间 Is Null) And U.用户名=user"
        .Open strSQL, gcnOracle, adOpenKeyset
                
        If .RecordCount <> 0 Then
            glngUserId = .Fields("ID").Value                '当前用户id
            gstrUserCode = .Fields("编号").Value            '当前用户编码
            gstrUserName = .Fields("姓名").Value            '当前用户姓名
            gstrUserAbbr = IIf(IsNull(.Fields("简码").Value), "", .Fields("简码").Value)          '当前用户简码
            glngDeptId = .Fields("部门id").Value            '当前用户部门id
            gstrDeptCode = .Fields("部门编码").Value        '当前用户
            gstrDeptName = .Fields("部门名称").Value        '当前用户
        Else
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
    Err = 0
End Sub

Public Function ChooseIME(cmbIME As Object) As Boolean
    Dim varIME As Variant
    Dim i As Integer
    Dim strIme As String
    
    varIME = OS.SystemImes
    If Not IsArray(varIME) Then
        MsgBox "你还没安装任何汉字输入法，不能使用本功能。" & vbCrLf & _
               "输入法的安装可在控制面板中完成。", vbInformation, gstrSysName
        Exit Function
    End If
    cmbIME.Clear
    cmbIME.AddItem "不自动开启"
    strIme = zlDatabase.GetPara("输入法")
    For i = LBound(varIME) To UBound(varIME)
        cmbIME.AddItem varIME(i)
        If strIme = varIME(i) Then cmbIME.ListIndex = i + 1
    Next
    If cmbIME.ListIndex < 0 Then cmbIME.ListIndex = 0
    ChooseIME = True
End Function

Public Function IsCheckConstraint(ByVal strOwner As String, ByVal strTableName As String, ByVal strColumnName As String, ByVal bytType As Byte) As Boolean
'获取Check约束内容
'bytType
'  1: 是否为 Check In (0,1) 约束
'  2: 是否为 Check Is Not Null 约束
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo ErrH
    strTmp = "Select A.Search_Condition from All_Constraints A, All_Cons_Columns B " _
           & "Where A.Constraint_Name = B.Constraint_Name and A.owner=[1] and a.Table_Name=[2] and B.Column_Name=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "", strOwner, strTableName, strColumnName)
    If Not rsTmp.EOF And IsNull(rsTmp!search_condition) = False Then
        Select Case bytType
            Case 1: If InStr(rsTmp!search_condition, "(0,1)") Or InStr(rsTmp!search_condition, "(1,0)") Then IsCheckConstraint = True
            Case 2: If InStr(UCase(rsTmp!search_condition), "IS NOT NULL") Or InStr(UCase(rsTmp!search_condition), "IS NULL") And InStr(UCase(rsTmp!search_condition), "NOT") Then IsCheckConstraint = True
        End Select
    End If
    rsTmp.Close
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function IsPathProperty(strOwner As String, strTable As String) As String
'说明：读外键约束是否指向路径结果性质表
'返回：从表外键列名;主表列名;主表名称
    Dim i As Integer
    Dim bln编码 As Boolean, blnID As Boolean, bln名称 As Boolean, bln外键 As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    IsPathProperty = ";"
    On Error GoTo errHandle
    
    Set rsTmp = zlDatabase.OpenSQLRecord("select * from " & strOwner & "." & strTable & " where rownum=0", "")
    If rsTmp Is Nothing Then Exit Function
    
    For i = 0 To rsTmp.Fields.Count - 1
        If rsTmp.Fields(i).Name = "编码" Then
            bln编码 = True
        ElseIf rsTmp.Fields(i).Name = "ID" Then
            blnID = True
        ElseIf rsTmp.Fields(i).Name = "名称" Then
            bln名称 = True
        End If
    Next
    rsTmp.Close
    If ((blnID Or bln编码) And bln名称) = False Then Exit Function
    
    strTmp = "Select b.Column_Name, c.Column_Name r_column_name,c.TABLE_NAME r_table_name " _
           & "From All_Constraints A, All_Cons_Columns B, All_Cons_Columns C " _
           & "Where a.Constraint_Name = b.Constraint_Name And a.r_Constraint_Name = c.Constraint_Name And a.Constraint_Type = 'R' " _
           & "  And a.owner=[1] and a.table_name=[2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "获取主从表字段外键名称", strOwner, strTable)
    Do While rsTmp.EOF = False
        If UCase(Nvl(rsTmp!column_name)) = "资源ID" And UCase(Nvl(rsTmp!r_table_name)) = "RESOURCEINFO" Then
            '此类条件为BH环境，排除在外
            IsPathProperty = ";;RESOURCEINFO"
        Else
            IsPathProperty = Nvl(rsTmp!column_name) & ";" & Nvl(rsTmp!r_column_name) & ";" & Nvl(rsTmp!r_table_name)
            Exit Do
        End If
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
