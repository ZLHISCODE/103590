VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClearData 
   BackColor       =   &H80000005&
   Caption         =   "数据清除"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmClearData.frx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   7500
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClear 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   3060
      TabIndex        =   10
      Top             =   4770
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全选(&C)"
      Height          =   350
      Left            =   1890
      TabIndex        =   9
      Top             =   4770
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwTable 
      Height          =   3135
      Left            =   1920
      TabIndex        =   8
      Top             =   1500
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "表名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "参照表"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "执行(&E)…"
      Height          =   350
      Left            =   5640
      TabIndex        =   6
      Top             =   4770
      Width           =   1100
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1920
      MaxLength       =   256
      TabIndex        =   3
      Top             =   1065
      Width           =   4485
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "…"
      Height          =   300
      Left            =   6420
      TabIndex        =   2
      Top             =   1050
      Width           =   300
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   690
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog cmmFile 
      Left            =   5070
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "清除表数据(&T)"
      Height          =   180
      Index           =   2
      Left            =   690
      TabIndex        =   7
      Top             =   1620
      Width           =   1170
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "安装文件(&F)"
      Height          =   180
      Index           =   0
      Left            =   870
      TabIndex        =   5
      Top             =   1110
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用系统(&S)"
      Height          =   180
      Index           =   1
      Left            =   870
      TabIndex        =   4
      Top             =   750
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmClearData.frx":04F9
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据清除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   960
   End
End
Attribute VB_Name = "frmClearData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsSystem As New ADODB.Recordset
Dim mrsConstraint As New ADODB.Recordset '保存着当前系统所有外键约束
Dim mstr所有者 As String '保存当前系统的所有者名
Dim mcolTable As New Collection '当前系统固定的表

Private Sub cmdClear_Click()
    Dim i As Integer
    For i = 1 To lvwTable.ListItems.Count
        lvwTable.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdExecute_Click()
    Dim colTemp As New Collection    '要删除的表记录
    Dim strErr As String             '记录异常的表
    Dim strDelete As String          '记录已经删除的表
    Dim strLoop As String            '上一个循环周期所删除的表
    Dim lngCount As Long
    Dim lst As ListItem
    Dim strTable As String
    Dim blnDelete As Boolean
    Dim strRemarks As String
    Dim strNote As String
    
    '得到要删除的表
    For Each lst In lvwTable.ListItems
        If lst.Checked = True Then
            colTemp.Add lst.Text, lst.Key
        End If
    Next
    If colTemp.Count = 0 Then Exit Sub
    If MsgBox("本操作是非常危险的，你确认已经正确选择了应该删除的表吗？", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '验证身份并输入操作说明
    If Not CheckAuditStatus("0206", "执行", strRemarks) Then Exit Sub
    frmMDIMain.Enabled = False
    Enabled = False
    frmWait.BeginWait "正在删除所选表的数据……"
    On Error Resume Next
    lngCount = 1
    Do While lngCount <= colTemp.Count
        '得到该表对应的列表记录
        strTable = colTemp(lngCount)           '表名用变量保存，效率更高
        
        blnDelete = CanDelete(strTable, strDelete) '对每一个表，缺省认为它是不能删除的
        If blnDelete = True Then
            '该表没有什么参照的，可以直接删除
            frmMDIMain.stbThis.Panels(2).Text = "现在删除" & strTable & "的数据……"
            gcnOracle.Execute "delete from " & mstr所有者 & "." & strTable
            strNote = strNote & "," & strTable
            If Err <> 0 Then
                Debug.Print Err.Description
                Err.Clear
                strErr = strErr & strTable & vbCrLf
            End If
            '不管出现错误与否，都把该表从集合中删除，以免死循环
            colTemp.Remove lngCount
            strDelete = strDelete & "[" & strTable & "]"
        Else
            '判断下一表
            lngCount = lngCount + 1
        End If
        '如果已经到末尾了，就重头开始
        If lngCount > colTemp.Count Then
            If strDelete = strLoop Then
                '循环了一圈，一个表都没删除，说明数据有问题
                If Right(strDelete, 1) = "," Then strDelete = Mid(strDelete, 1, Len(strDelete) - 1) '去掉最后的逗号
                MsgBox "你的数据表不完整，部分表不能删除。" & IIf(strDelete = "", "", "但是下列表已被删除：" & vbCrLf & vbCrLf & strDelete), vbExclamation, gstrSysName
                frmMDIMain.stbThis.Panels(2).Text = ""
                frmWait.EndWait
                Enabled = True
                frmMDIMain.Enabled = True
                '插入重要操作日志
                If strNote <> "" Then
                    Call SaveAuditLog(3, "执行", "成功将“" & Split(cmbSystem.Text, " ")(0) & "”中的数据表“" & Mid(strNote, 2) & "”清空", strRemarks)
                End If
                Exit Sub
            End If
            lngCount = 1
            strLoop = strDelete
        End If
        
    Loop
    frmMDIMain.stbThis.Panels(2).Text = ""
    frmWait.EndWait
    Enabled = True
    frmMDIMain.Enabled = True
    If Right(strErr, 1) = "," Then strErr = Mid(strErr, 1, Len(strErr) - 1) '去掉最后的逗号
    MsgBox "数据删除操作执行完毕。" & IIf(strErr = "", "", "但是下列表未正常删除：" & vbCrLf & vbCrLf & strErr), vbExclamation, gstrSysName
    '插入重要操作日志
    If strNote <> "" Then
        Call SaveAuditLog(3, "执行", "成功将“" & Split(cmbSystem.Text, " ")(0) & "”中的数据表“" & Mid(strNote, 2) & "”清空", strRemarks)
    End If
End Sub

Private Function CanDelete(ByVal strTable As String, strDelete As String) As Boolean
    Dim lst As ListItem
    Dim strTemp As String
    Dim varRefTable As Variant
    Dim i As Long
    
    CanDelete = False
    Set lst = lvwTable.ListItems("C" & strTable)
    If lst.SubItems(1) = "" Then
        CanDelete = True
    Else
        varRefTable = Split(lst.SubItems(1), ",")
        For i = LBound(varRefTable) To UBound(varRefTable)
            If varRefTable(i) <> strTable And varRefTable(i) <> strTable & "(*)" Then
                '自己不用判断
                If InStr(varRefTable(i), "(*)") = 0 Then
                    '得到一个有约束的参照表，判断它是否已经删除
                    If InStr(strDelete, "[" & varRefTable(i) & "]") = 0 Then
                        '有一参照表没删除，就不能删除本表
                        Exit For
                    End If
                Else
                    strTemp = Mid(varRefTable(i), 1, InStr(varRefTable(i), "(*)") - 1)
                    If CanDelete(strTemp, strDelete) = False Then
                        '有一参照表没删除，就不能删除本表
                        Exit For
                    End If
                End If
            End If
        Next
        If i > UBound(varRefTable) Then CanDelete = True
    End If
    
End Function

Private Sub cmdFile_Click()
    Dim lst As ListItem
    Dim strTemp As String

    cmmFile.Filter = "应用安装配置文件|zlSetup.ini"
    cmmFile.FileName = txtFile.Text
    cmmFile.ShowOpen
    If cmmFile.FileName = "" Then Exit Sub
    '就是以前选定的文件
    If txtFile.Text = cmmFile.FileName Then Exit Sub
    '重新进行检查
    txtFile.Text = cmmFile.FileName
    If CheckIniFile(txtFile.Text, False) = False Then
        cmdExecute.Enabled = False
        lvwTable.ListItems.Clear
    Else
        cmdExecute.Enabled = True
        Call FillTable
    End If
End Sub

Private Sub cmbSystem_Click()
    Dim rsTemp As New ADODB.Recordset
On Error GoTo ErrHandle
    
    '与上次相同，不需要招行
    If mrsSystem.Filter = "编号=" & cmbSystem.ItemData(cmbSystem.ListIndex) Then Exit Sub
    
    MousePointer = 11
    mrsSystem.Filter = "编号=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    If mrsSystem.RecordCount = 0 Then
        cmdExecute.Enabled = False
        txtFile.Text = ""
    Else
        cmdExecute.Enabled = True
        mstr所有者 = mrsSystem("所有者")
        '读出系统安装脚本的位置
        Dim varOut As Variant
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Loadandunload.Get_Sysfile_name", Val(cmbSystem.ItemData(cmbSystem.ListIndex)), 1)
        'txtFile.Text = IIf(IsNull(varOut(0)), "", varOut(0))
        If rsTemp.RecordCount <= 0 Then
            txtFile.Text = ""
        Else
            txtFile.Text = IIf(IsNull(rsTemp("文件名")), "", rsTemp("文件名"))
        End If
        rsTemp.Close
        
        '读出当前系统的所有约束
        If mrsConstraint.State = 1 Then mrsConstraint.Close
        gstrSQL = "select A.table_name ,B.table_name r_table_name,A.DELETE_RULE" & _
                   " from all_constraints A,all_constraints b" & _
                   " where A.owner='" & mstr所有者 & "' AND b.OWNER='" & mstr所有者 & "' and A.r_owner=B.owner" & _
                   "     and A.R_CONSTRAINT_NAME=b.constraint_name And Instr(A.Table_NAME,'BIN$')<=0"
        mrsConstraint.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    End If
    '对文件的检查
    If CheckIniFile(txtFile.Text, True) = False Then
        cmdExecute.Enabled = False
        lvwTable.ListItems.Clear
    Else
        cmdExecute.Enabled = True
        Call FillTable
    End If
    MousePointer = 0
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub FillTable()
    Dim rsTemp As New ADODB.Recordset
    Dim strTable As String
    Dim strTemp As String
    Dim lst As ListItem
On Error GoTo ErrHandle
    
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "select table_name from all_tables A where owner='" & mstr所有者 & "' And Instr(A.Table_NAME,'BIN$')<=0 order by table_name", gcnOracle, adOpenStatic, adLockReadOnly
    
    lvwTable.ListItems.Clear
    If rsTemp.RecordCount = 0 Then
        cmdExecute.Enabled = False
        Exit Sub
    End If
    '装入表格
    On Error Resume Next
    Do Until rsTemp.EOF
        strTable = rsTemp("TABLE_NAME")
        '得到它是否是基础表
        strTemp = mcolTable("C" & strTable)
        If Err <> 0 Then
            '不是，所以加入
            Err.Clear
            Set lst = lvwTable.ListItems.Add(, "C" & strTable, strTable)
            '得到它的所有下级表
            mrsConstraint.Filter = "r_table_name='" & strTable & "'"
            strTemp = ""
            Do Until mrsConstraint.EOF
                strTemp = strTemp & mrsConstraint("TABLE_NAME") & _
                    IIf(mrsConstraint("DELETE_RULE") = "NO ACTION", "", "(*)") & ","
                mrsConstraint.MoveNext
            Loop
            If strTemp <> "" Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1) '去掉最后一个的逗号
            lst.SubItems(1) = strTemp
        End If
        rsTemp.MoveNext
    Loop

    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdSelect_Click()
    Dim i As Integer
    For i = 1 To lvwTable.ListItems.Count
        lvwTable.ListItems(i).Checked = True
    Next
End Sub

Private Sub Command1_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub Form_Resize()
    Dim sngTemp As Single
    
    On Error Resume Next
    sngTemp = IIf(ScaleWidth > 6000, ScaleWidth, 6000)
    cmbSystem.Width = sngTemp - cmbSystem.Left - 200
    cmdFile.Left = sngTemp - cmdFile.Width - 200
    txtFile.Width = cmdFile.Left - 15 - txtFile.Left
    lvwTable.Width = cmbSystem.Width
    cmdExecute.Left = lvwTable.Left + lvwTable.Width - cmdExecute.Width
    
    sngTemp = IIf(ScaleHeight > 3000, ScaleHeight, 3000)
    cmdExecute.Top = sngTemp - cmdExecute.Height - 200
    cmdClear.Top = cmdExecute.Top
    cmdSelect.Top = cmdExecute.Top
    lvwTable.Height = cmdExecute.Top - lvwTable.Top - 100
'    lbl说明.Width = ScaleWidth - 200 - lbl说明.Left
'    lbl说明.Height = ScaleHeight - 200 - lbl说明.Top
    
End Sub

Private Function CheckIniFile(FileName As String, blnCmb As Boolean) As Boolean
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strIniPath As String
    Dim strTemp As String
    Dim strTable As String
    Dim lngCount As Long
    On Error Resume Next
    
    strIniPath = Mid(FileName, 1, Len(FileName) - 11)
    '相关文件匹配性检查
    If Dir(strIniPath & "zlAppData.sql") = "" Then
        MsgBox "应用数据文件" & strIniPath & "zlAppData.sql丢失，不能继续。", vbExclamation, gstrSysName
        txtFile.Text = ""
        Exit Function
    End If
    
    If mrsSystem.EOF Then
        txtFile.Text = ""
        Exit Function
    End If
    '配置文件正确性检查
    Set objText = objFile.OpenTextFile(FileName)
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[系统号]" Then
        If Val(Mid(strTemp, 6)) <> mrsSystem("编号") \ 100 Then
            If blnCmb = False Then MsgBox "所选文件不是该系统的安装配置文件", vbExclamation, gstrSysName
            txtFile.Text = ""
            Exit Function
        End If
    Else
        Err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine) '取出系统名
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[版本号]" Then
        If InStr(1, mrsSystem("版本号"), Trim(Mid(strTemp, 6))) = 0 Then
            MsgBox "选择文件与该系统版本不符", vbExclamation, gstrSysName
            txtFile.Text = ""
            Exit Function
        End If
    Else
        Err.Raise 10
    End If
    
    objText.Close
    
    If Err <> 0 Then
        CheckIniFile = False
        MsgBox "安装配置文件不正确" & vbNewLine & Err.Description, vbExclamation, gstrSysName
        txtFile.Text = ""
        Exit Function
    End If
    '得到基础表
    '清空集合
    For lngCount = 1 To mcolTable.Count
        mcolTable.Remove 1
    Next
    '增加所有的基础表
    Set objText = objFile.OpenTextFile(strIniPath & "zlAppData.sql")
    Do Until objText.AtEndOfStream
        strTemp = UCase(objText.ReadLine())
        lngCount = InStr(strTemp, "INTO")
        If lngCount > 0 Then '去掉前面的"insert into"
            strTemp = Mid(strTemp, lngCount + 4)
            lngCount = InStr(strTemp, "(") '去掉后面的"("
            If lngCount > 0 Then
                strTemp = Trim(Mid(strTemp, 1, lngCount - 1))
                If strTemp <> "" And strTemp <> strTable Then
                    strTable = strTemp
                    mcolTable.Add strTable, "C" & strTable
                End If
            End If
        End If
    Loop
    objText.Close
    CheckIniFile = True
End Function

Private Sub Form_Load()
    
On Error GoTo ErrHandle
    frmMDIMain.MousePointer = 11

    '显示所有可显示的系统
    
    If gblnDBA = True Then
        Set mrsSystem = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set mrsSystem = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", UCase(gstrUserName))
    End If
    
    Do Until mrsSystem.EOF
        cmbSystem.AddItem mrsSystem("名称") & " v" & mrsSystem("版本号") & "（" & mrsSystem("编号") & "）"
        cmbSystem.ItemData(cmbSystem.NewIndex) = mrsSystem("编号")
        mrsSystem.MoveNext
    Loop
    If mrsSystem.RecordCount > 0 Then
        cmbSystem.ListIndex = 0
    Else
        cmdExecute.Enabled = False
    End If
    frmMDIMain.MousePointer = 0
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsSystem.State = 1 Then mrsSystem.Close
    Set mrsSystem = Nothing
    Set mcolTable = Nothing
    mstr所有者 = ""
End Sub

Private Sub lvwTable_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'
    SetCheck Item
End Sub


Private Sub SetCheck(ByVal Item As MSComctlLib.ListItem)
    Dim varTable As Variant
    Dim strTable As String
    Dim i As Integer
    Dim lst As ListItem
    
On Error GoTo ErrHandle
    If Item.Checked = False Then
        '如果子级的表没有删除，那其依赖它的表也不能删除
        mrsConstraint.Filter = "table_name='" & Item.Text & "'"
        strTable = ""
        Do Until mrsConstraint.EOF
            If mrsConstraint("DELETE_RULE") = "NO ACTION" Then
                '级联删除的就不考虑了
                strTable = strTable & mrsConstraint("R_TABLE_NAME") & ","
            End If
            mrsConstraint.MoveNext
        Loop
        If strTable <> "" Then strTable = Mid(strTable, 1, Len(strTable) - 1) '去掉最后一个的逗号
        varTable = Split(strTable, ",")
        For i = LBound(varTable) To UBound(varTable)
            If varTable(i) <> Item.Text Then
                Set lst = lvwTable.ListItems("C" & varTable(i))
                lst.Checked = False
                '进行递归调用
                SetCheck lst
            End If
        Next
    Else
        varTable = Split(Item.SubItems(1), ",")
        For i = LBound(varTable) To UBound(varTable)
            If InStr(varTable(i), "(*)") = 0 And varTable(i) <> Item.Text Then
                Set lst = lvwTable.ListItems("C" & varTable(i))
                lst.Checked = True
                '进行递归调用
                SetCheck lst
            End If
        Next
    
    End If
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub


Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub

