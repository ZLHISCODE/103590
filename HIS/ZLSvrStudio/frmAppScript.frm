VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppScript 
   BackColor       =   &H80000005&
   Caption         =   "置换安装脚本"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmAppScript.frx":0000
   ScaleHeight     =   5610
   ScaleWidth      =   6030
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   975
      TabIndex        =   11
      Top             =   1665
      Width           =   4320
      Begin VB.TextBox txtGroup 
         Height          =   300
         Left            =   465
         TabIndex        =   13
         Top             =   1065
         Width           =   3630
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   300
         Left            =   465
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   435
         Width           =   3630
      End
      Begin VB.PictureBox picXp 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   210
         ScaleHeight     =   1125
         ScaleWidth      =   3975
         TabIndex        =   14
         Top             =   135
         Width           =   3975
         Begin VB.OptionButton optOld 
            BackColor       =   &H80000005&
            Caption         =   "替换原独立可选数据组(&E)"
            Height          =   285
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   2520
         End
         Begin VB.OptionButton optNew 
            BackColor       =   &H80000005&
            Caption         =   "生成新可选数据组(&N)"
            Height          =   285
            Left            =   0
            TabIndex        =   15
            Top             =   615
            Value           =   -1  'True
            Width           =   2055
         End
      End
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "执行(&E)…"
      Height          =   350
      Left            =   975
      TabIndex        =   10
      Top             =   3210
      Width           =   1100
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   5970
      TabIndex        =   7
      Top             =   5070
      Visible         =   0   'False
      Width           =   6030
      Begin MSComctlLib.ProgressBar pgbState 
         Height          =   180
         Left            =   135
         TabIndex        =   8
         Top             =   255
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正在检查"
         Height          =   180
         Left            =   135
         TabIndex        =   9
         Top             =   60
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdGetIni 
      Caption         =   "选择(&S)…"
      Height          =   350
      Left            =   4215
      TabIndex        =   3
      Top             =   1020
      Width           =   1095
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   645
      Width           =   3570
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   1110
      Left            =   975
      TabIndex        =   6
      Top             =   3750
      Width           =   4395
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   960
      TabIndex        =   5
      Top             =   1380
      Width           =   4350
   End
   Begin VB.Label lblFileCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "安装配置文件"
      Height          =   180
      Left            =   960
      TabIndex        =   4
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用系统"
      Height          =   180
      Left            =   975
      TabIndex        =   1
      Top             =   705
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "置换安装脚本"
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
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmAppScript.frx":04F9
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmAppScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strIniPath    As String                 '安装配置文件目录
Dim intDefSysCode As Integer                '系统编号
Dim strDefSysName As String                 '系统名称
Dim strDefVersion As String                 '版本号
Dim strDefSpace   As String                 '表空间
Dim strDefUser    As String                 '新的缺省用户名
Dim strDefData    As String                 '用户可选的数据
Dim aryRow() As String
Dim aryVal() As String

Dim objFile As New FileSystemObject
Dim objText As TextStream

Dim rsTemp As New ADODB.Recordset
Dim rsObjects As New ADODB.Recordset
Dim rsColumns As New ADODB.Recordset
Dim strSQL As String, strTemp As String
Dim intCount As Integer

Private Sub cmbSystem_Click()

    cmbSystem.Tag = GetOwnerName(Val(cmbSystem.ItemData(cmbSystem.ListIndex)), gcnOracle)
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Loadandunload.Get_Sysfile_name", Val(cmbSystem.ItemData(cmbSystem.ListIndex)), 1)
    With rsTemp
        If Not .EOF And Not .BOF Then
            If gobjFile.FileExists(.Fields(0).value) Then
                lblFileName.Caption = .Fields(0).value
            Else
                lblFileName.Caption = ""
            End If
        End If
    End With
    If CheckIniFile(lblFileName.Caption, False) = False Then
        fraGroup.Enabled = False
        cmdGen.Enabled = False
        lblFileName.Caption = ""
    Else
        fraGroup.Enabled = True
        cmdGen.Enabled = True
    End If

End Sub

Private Sub cmdGen_Click()
    
    Dim strWriteFile As String
    Dim bytTreeData As Byte     '是否树性型结构
    
    If optOld.value = True Then
        If MsgBox("是否替换原数据组“" & cboGroup.List(cboGroup.ListIndex) & "”？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        strWriteFile = strIniPath & "zlSelData" & cboGroup.ItemData(cboGroup.ListIndex) & ".sql"
    ElseIf cboGroup.Tag >= 9 Then
        If MsgBox("由于现有数据组超过限制，只能产生不能在下次" & vbCr & "安装中直接生效的文件：" & strIniPath & "zlSelDataTemp.sql" & vbCr & vbCr & "继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        strWriteFile = strIniPath & "zlSelDataTemp.sql"
    Else
        strWriteFile = strIniPath & "zlSelData" & cboGroup.Tag + 1 & ".sql"
    End If
    
    If CheckIniFile(lblFileName.Caption, True) = False Then
        cmdGen.Enabled = False
        lblFileName.Caption = ""
        cmdGetIni.SetFocus
        Exit Sub
    End If
    
    err = 0
    On Error Resume Next
    Kill strWriteFile
    err = 0
    Open strWriteFile For Binary Access Write As #1
    If err <> 0 Then
        MsgBox "由于不能打开数据组文件，不能继续。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    picStatus.Visible = True
    Enabled = False
    
    Dim strTables As String, strInsert As String, strValues As String
    Dim blnFather As Boolean
    
    strTables = ""
    With rsTemp
        If gblnDBA Then
            strSQL = "select TABLE_NAME from DBA_TABLES where OWNER='" & cmbSystem.Tag & "'"
        Else
            strSQL = "select TABLE_NAME from USER_TABLES"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        Do While Not .EOF
            strTables = strTables & "," & !Table_Name
            .MoveNext
        Loop
    End With
    strTables = strTables & ","
    
    Do While strTables <> ","
        aryRow = Split(strTables, ",")
        For intCount = 1 To UBound(aryRow) - 1
            If InStr(1, strTables, "," & aryRow(intCount) & ",") > 0 Then
                blnFather = True
                If gblnDBA Then
                    strSQL = "Select TABLE_NAME" & _
                            " From (Select Table_Name, Constraint_Name From Dba_Constraints Where Owner = '" & cmbSystem.Tag & "' and table_name<>'" & aryRow(intCount) & "') t," & _
                            "      (select distinct r_constraint_name" & _
                            "       from dba_constraints" & _
                            "       where OWNER='" & cmbSystem.Tag & "' and constraint_type='R' and table_name='" & aryRow(intCount) & "') r" & _
                            " Where t.Constraint_Name = r.r_Constraint_Name"
                Else
                    strSQL = "Select TABLE_NAME" & _
                            " From (Select Table_Name, Constraint_Name From Dba_Constraints Where table_name<>'" & aryRow(intCount) & "') t," & _
                            "      (select distinct r_constraint_name" & _
                            "       from dba_constraints" & _
                            "       where constraint_type='R' and table_name='" & aryRow(intCount) & "') r" & _
                            " Where t.Constraint_Name = r.r_Constraint_Name"
                End If
                With rsTemp
                    If .State = adStateOpen Then .Close
                    .Open strSQL, gcnOracle, adOpenKeyset
                    Do While Not .EOF
                        If InStr(1, strTables, "," & !Table_Name & ",") > 0 Then
                            blnFather = False
                            Exit Do
                        End If
                        .MoveNext
                    Loop
                End With
                If blnFather Then
                    aryVal = Split(strTables, "," & aryRow(intCount) & ",")
                    strTables = aryVal(0) & "," & aryVal(1)
                    lblStatus.Caption = aryRow(intCount)
                    With rsTemp
                        If gblnDBA Then
                            strSQL = "SELECT COLUMN_NAME,DATA_TYPE" & _
                                    " From DBA_TAB_COLUMNS" & _
                                    " WHERE OWNER='" & cmbSystem.Tag & "' and TABLE_NAME='" & aryRow(intCount) & "'"
                        Else
                            strSQL = "SELECT COLUMN_NAME,DATA_TYPE" & _
                                    " From USER_TAB_COLUMNS" & _
                                    " WHERE TABLE_NAME='" & aryRow(intCount) & "'"
                        End If
                        If .State = adStateOpen Then .Close
                        .Open strSQL, gcnOracle, adOpenKeyset
                        
                        strInsert = ""
                        strValues = ""
                        bytTreeData = 0
                        Do While Not .EOF
                            If !COLUMN_NAME = "ID" Then bytTreeData = bytTreeData + 1
                            If !COLUMN_NAME = "上级ID" Then bytTreeData = bytTreeData + 1
                            Select Case !DATA_TYPE
                            Case "NUMBER", "INTEGER"
                                strInsert = strInsert & "," & !COLUMN_NAME
                                strValues = strValues & "||','||decode(" & !COLUMN_NAME & ",null,'null'," & !COLUMN_NAME & ")"
                            Case "VARCHAR2"
                                strInsert = strInsert & "," & !COLUMN_NAME
                                strValues = strValues & "||','||chr(39)||" & !COLUMN_NAME & "||chr(39)"
                            Case "DATE"
                                strInsert = strInsert & "," & !COLUMN_NAME
                                strValues = strValues & "||','||decode(" & !COLUMN_NAME & ",null,'null','to_date('||chr(39)||to_char(" & !COLUMN_NAME & ",'YYYY-MM-DD HH24:MI:SS')||chr(39)||','||chr(39)||'YYYY-MM-DD HH24:MI:SS'||chr(39)||')')"
                            Case Else
                            End Select
                            .MoveNext
                        Loop
                        If strInsert <> "" Then
                            strSQL = "select " & "'insert into " & aryRow(intCount) & "(" & Mid(strInsert, 2) & ")" & " values(" & Mid(strValues, 5) & "||');'" & _
                                    " from " & IIf(gblnDBA, cmbSystem.Tag & ".", "") & aryRow(intCount)
                            If bytTreeData = 2 Then
                                strSQL = strSQL & " start with 上级ID is null connect by prior ID=上级ID order by level"
                            End If
                            If .State = adStateOpen Then .Close
                            .Open strSQL, gcnOracle, adOpenKeyset
                            If Not .EOF Then
                                Put #1, , "--" & aryRow(intCount) & vbCrLf
                                strSQL = "delete from " & aryRow(intCount) & ";"
                                Put #1, , strSQL & vbCrLf
                            End If
                            Do While Not .EOF
                                strSQL = Trim(.Fields(0).value)
                                Put #1, , strSQL & vbCrLf
                                .MoveNext
                            Loop
                        End If
                    End With
                End If
            End If
        Next
    Loop
    Close #1
    
    If optOld.value = False And cboGroup.Tag < 9 Then
        Set objText = objFile.OpenTextFile(lblFileName.Caption)
        strTemp = Trim(objText.ReadAll)
        objText.Close
        
        err = 0
        Open lblFileName.Caption For Binary Access Write As #1
        If err <> 0 Then
            strSQL = "由于不能打开安装配置件，请检查后重新运行；" & _
                    vbCr & "或者手工在文件" & lblFileName.Caption & "的 [数据组] 行末，" & _
                    vbCr & "增加“||" & txtGroup.Text & "完成。"
            MsgBox strSQL, vbExclamation, gstrSysName
            Exit Sub
        End If
        
        aryRow = Split(strTemp, vbCrLf)
        For intCount = 0 To UBound(aryRow)
            If Left(aryRow(intCount), 5) = "[数据组]" Then
                aryRow(intCount) = aryRow(intCount) & "||" & txtGroup.Text
            End If
            Put #1, , aryRow(intCount) & vbCrLf
        Next
        Close #1
    End If
    
    picStatus.Visible = False
    Enabled = True
    MsgBox "数据安装脚本处理完毕！", vbExclamation, gstrSysName
    
End Sub

Private Sub cmdGetIni_Click()
    With frmMDIMain.DlgMain
        .FileName = lblFileName.Caption
        .DialogTitle = "选择应用安装配置文件"
        .Filter = "(应用安装配置文件)|zlSetup.ini"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            lblFileName.Caption = .FileName
        End If
    End With
    
    If CheckIniFile(lblFileName.Caption, True) = False Then
        fraGroup.Enabled = False
        cmdGen.Enabled = False
        lblFileName.Caption = ""
        cmdGetIni.SetFocus
    Else
        fraGroup.Enabled = True
        cmdGen.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    
    lblMain.Caption = "如果当前系统的数据具有代表性，而希望在其他合法授权用户安装过程中能直接使用该数据进行安装，可以使用该程序产生新的安装脚本文件。" & _
        vbCrLf & vbCrLf & "但对于包含在数据表的大对象(LOB)和LONG型列，系统无法产生有效的安装脚本。"
    
    On Error GoTo ErrHandle
    txtGroup.Text = gobjRegister.zlRegInfo("单位名称") & "数据"
    
    If gblnDBA Then
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", gstrUserName)
    End If
    Do While Not rsTemp.EOF
        cmbSystem.AddItem rsTemp!名称 & " v" & rsTemp!版本号 & "（" & rsTemp!编号 & "）"
        cmbSystem.ItemData(cmbSystem.NewIndex) = rsTemp!编号
        rsTemp.MoveNext
    Loop
    If cmbSystem.ListCount = 0 Then
        cmdGetIni.Enabled = False
        cmdGen.Enabled = False
    End If
    If cmbSystem.ListCount > 0 Then cmbSystem.ListIndex = 0
    If cmbSystem.ListCount = 1 Then cmbSystem.Locked = True
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim sngWidth As Long '最小宽度
    
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
        
    With lblMain
        .Top = cmdGen.Top + cmdGen.Height + 200
        .Height = ScaleHeight - picStatus.Height - .Top - 100
        .Left = lblFileName.Left
        .Width = ScaleWidth - .Left - imgMain.Left
    End With
    
    sngWidth = IIf(ScaleWidth < 5600, 5600, ScaleWidth)
    cmbSystem.Width = sngWidth - cmbSystem.Left - 300
    cmdGetIni.Left = cmbSystem.Left + cmbSystem.Width - cmdGetIni.Width
    lblFileName.Width = sngWidth - lblFileName.Left - 300
    fraGroup.Width = sngWidth - fraGroup.Left - 300
    cboGroup.Width = fraGroup.Width - 600
    txtGroup.Width = cboGroup.Width
    
End Sub

Private Function CheckIniFile(FileName As String, Optional blnMsg As Boolean) As Boolean
    err = 0
    On Error Resume Next
    
    strIniPath = Mid(FileName, 1, Len(FileName) - 11)
    '相关文件匹配性检查
    strTemp = ""
    If Dir(strIniPath & "zlSequence.sql") = "" Then strTemp = strTemp & vbCr & "序列文件" & strIniPath & "zlSequence.sql"
    If Dir(strIniPath & "zlTable.sql") = "" Then strTemp = strTemp & vbCr & "数据表文件" & strIniPath & "zlTable.sql"
    If Dir(strIniPath & "zlConstraint.sql") = "" Then strTemp = strTemp & vbCr & "约束文件" & strIniPath & "zlConstraint.sql"
    If Dir(strIniPath & "zlIndex.sql") = "" Then strTemp = strTemp & vbCr & "索引文件" & strIniPath & "zlIndex.sql"
    If Dir(strIniPath & "zlView.sql") = "" Then strTemp = strTemp & vbCr & "视图文件" & strIniPath & "zlView.sql"
    If Dir(strIniPath & "zlProgram.sql") = "" Then strTemp = strTemp & vbCr & "函数过程文件" & strIniPath & "zlProgram.sql"
    
    '不检查,因为9系统没有此文件
    'If Dir(strIniPath & "zlPackage.sql") = "" Then strTemp = strTemp & vbCr & "包文件" & strIniPath & "zlPackage.sql"
    
    If Dir(strIniPath & "zlManData.sql") = "" Then strTemp = strTemp & vbCr & "管理数据文件" & strIniPath & "zlManData.sql"
    If Dir(strIniPath & "zlAppData.sql") = "" Then strTemp = strTemp & vbCr & "应用数据文件" & strIniPath & "zlAppData.sql"
    If strTemp <> "" Then
        If blnMsg Then MsgBox "以下服务器安装的相关文件丢失，不能继续，包括：" & strTemp, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '配置文件正确性检查
    Set objText = objFile.OpenTextFile(FileName)
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[系统号]" Then
        intDefSysCode = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[系统名]" Then
        strDefSysName = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[版本号]" Then
        strDefVersion = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[表空间]" Then
        strDefSpace = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[用户名]" Then
        strDefUser = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[数据组]" Then
        strDefData = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    
    objText.Close
    
    If err <> 0 Then
        CheckIniFile = False
        If blnMsg Then MsgBox "安装配置文件不正确", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '配置文件符合性检查
    If intDefSysCode <> cmbSystem.ItemData(cmbSystem.ListIndex) \ 100 Then
        If blnMsg Then MsgBox "选择文件不是该系统的安装配置文件。", vbExclamation, gstrSysName
        CheckIniFile = False
        Exit Function
    ElseIf InStr(1, cmbSystem.Text, Trim(strDefVersion)) = 0 Then
        If blnMsg Then MsgBox "选择文件与该系统版本不符。", vbExclamation, gstrSysName
        CheckIniFile = False
        Exit Function
    End If
    
    '数据分组可选文件匹配性检查
    cboGroup.Clear
    err = 0
    aryRow = Split(strDefData, "||")
    For intCount = 0 To UBound(aryRow)
        If InStr(1, aryRow(intCount), ">") = 0 Then
            If InStr(1, aryRow(intCount), "=") = 0 Then
                cboGroup.AddItem aryRow(intCount)
            Else
                cboGroup.AddItem Trim(Left(aryRow(intCount), InStr(1, aryRow(intCount), "=") - 1))
            End If
            cboGroup.ItemData(cboGroup.NewIndex) = intCount
        End If
    Next
    cboGroup.Tag = UBound(aryRow)
    If cboGroup.ListCount = 0 Then
        optOld.value = False
        optOld.Enabled = False
        cboGroup.AddItem "该系统无独立数据组。"
    Else
        optOld.Enabled = True
        cboGroup.ListIndex = 0
    End If
    
    If err = 0 Then
        CheckIniFile = True
    Else
        CheckIniFile = False
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    If picStatus.Visible Then Cancel = 1
End Sub

Private Sub optNew_Click()
    cboGroup.Enabled = False
    txtGroup.Enabled = True
End Sub

Private Sub optOld_Click()
    cboGroup.Enabled = True
    txtGroup.Enabled = False
End Sub

Private Sub picStatus_Resize()
    pgbState.Width = picStatus.ScaleWidth - pgbState.Left * 2
End Sub


Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub

