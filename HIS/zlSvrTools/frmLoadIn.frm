VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoadIn 
   BackColor       =   &H80000005&
   Caption         =   "数据调入"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmLoadIn.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   8085
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "调入前清空表数据(&L)"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3690
      TabIndex        =   13
      Top             =   4590
      Value           =   1  'Checked
      Width           =   2085
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全清(&R)"
      Height          =   350
      Left            =   2310
      TabIndex        =   10
      Top             =   5040
      Width           =   1155
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   1020
      TabIndex        =   9
      Top             =   5040
      Width           =   1155
   End
   Begin VB.DriveListBox DriveBak 
      Height          =   300
      Left            =   3675
      TabIndex        =   6
      Top             =   1290
      Width           =   2880
   End
   Begin VB.DirListBox DirBak 
      Appearance      =   0  'Flat
      Height          =   2820
      Left            =   3675
      TabIndex        =   7
      Top             =   1620
      Width           =   2880
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "执行(&E)"
      Height          =   350
      Left            =   5460
      TabIndex        =   8
      Top             =   5040
      Width           =   1155
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5760
      Width           =   5625
   End
   Begin VB.ComboBox cmbSystem 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   4485
   End
   Begin MSComctlLib.ListView lvwTabs 
      Height          =   3675
      Left            =   1020
      TabIndex        =   4
      Top             =   1290
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   6482
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblTabs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据表选择(T)"
      Height          =   180
      Left            =   1020
      TabIndex        =   3
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件存放目录(&D)"
      Height          =   180
      Left            =   3690
      TabIndex        =   5
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对Loader程序了解的话，也可手工依次执行下面的命令："
      Height          =   180
      Index           =   3
      Left            =   1050
      TabIndex        =   11
      Top             =   5490
      Width           =   4500
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "调入系统(&S)"
      Height          =   180
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   780
      Width           =   990
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   240
      Picture         =   "frmLoadIn.frx":04F9
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据调入"
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
Attribute VB_Name = "frmLoadIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mrsTable As New ADODB.Recordset
Dim mstr所有者 As String '保存当前系统的所有者名
Dim mstrVer As String

Private Sub cmbSystem_Click()
    Call DirBak_Change
End Sub

Private Sub cmdSelect_Click()
    Dim i As Integer
    For i = 1 To lvwTabs.ListItems.Count
        lvwTabs.ListItems(i).Checked = True
    Next
    Call lvwTabs_ItemCheck(Nothing)
End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    For i = 1 To lvwTabs.ListItems.Count
        lvwTabs.ListItems(i).Checked = False
    Next
    Call lvwTabs_ItemCheck(Nothing)
End Sub

Private Sub DirBak_Change()
    Dim strFile As String
    Dim strPath As String
    Dim lst As ListItem
    
    strPath = DirBak.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    lvwTabs.ListItems.Clear
    txtFile.Text = ""
    strFile = UCase(Dir(strPath))
    Do Until strFile = ""
        If Right(strFile, 4) = ".LDR" Then
            mrsTable.Filter = "Table_name='" & Left(strFile, Len(strFile) - 4) & "'"
            If Not mrsTable.EOF Then
                Set lst = lvwTabs.ListItems.Add(, , mrsTable("TABLE_NAME"))
                lst.Checked = True
                '显示当前可执行的命令
                txtFile.Text = txtFile.Text & "SQLLDR" & mstrVer & " " & gstrUserName & "/" & gstrPassword & IIf(gstrServer = "", "", "@" & gstrServer) _
                    & " CONTROL=" & strPath & strFile & " DIRECT=TRUE" & vbCrLf
            End If
        End If
        strFile = UCase(Dir())
    Loop
    cmdExecute.Enabled = (txtFile.Text <> "")
End Sub

Private Sub lvwTabs_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim strFile As String
    Dim strPath As String
    Dim lst As ListItem
    
    strPath = DirBak.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    txtFile.Text = ""
    strFile = UCase(Dir(strPath))
    For Each lst In lvwTabs.ListItems
        If lst.Checked = True Then
            '显示当前可执行的命令
            txtFile.Text = txtFile.Text & "SQLLDR" & mstrVer & " " & gstrUserName & "/" & gstrPassword & IIf(gstrServer = "", "", "@" & gstrServer) _
                & " CONTROL=" & strPath & lst.Text & ".LDR DIRECT=TRUE" & vbCrLf
        End If
    Next
    cmdExecute.Enabled = (txtFile.Text <> "")
End Sub

Private Sub DriveBak_Change()
    Dim strDrive As String
        
    On Error GoTo UnDo
    strDrive = DirBak.Path
    DirBak.Path = DriveBak.Drive
    Exit Sub
UnDo:
    MsgBox "驱动器未准备好", vbExclamation, gstrSysName
    DriveBak.Drive = strDrive
End Sub

Private Sub cmdExecute_Click()
    Dim strPath As String
    Dim strErrTabs As String
    Dim strErrCons As String
    Dim strCommand As String
    Dim strTable As String
    Dim lngTemp As Long
    Dim lngProcess As Long
    Dim varTime As Variant
    Dim lst As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    strPath = DirBak.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    strErrTabs = ""
    strErrCons = ""
    
    If MsgBox("调入数据是一个很漫长的过程，且要破坏现有数据，" & vbCrLf & "你准备好了吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    SetEnable False
    frmWait.BeginWait "正在调入数据……"
    varTime = Now() '记录下开始导出的时间
    
    '逐个表进行loader
    For Each lst In lvwTabs.ListItems
        If lst.Checked Then
            strTable = lst.Text
            If chkDelete.value = 1 Then
                '将制向该表的外键约束disable
                gstrSQL = "select 'ALTER TABLE '||D.table_name||' DISABLE CONSTRAINT '||D.constraint_name" & _
                        " from user_constraints U,user_constraints D" & _
                        " where U.table_name='" & strTable & "' and U.constraint_type in('P','U')" & _
                        "       and U.constraint_name=D.r_constraint_name"
                With rsTemp
                    If .State = adStateOpen Then .Close
                    .Open gstrSQL, gcnOracle, adOpenKeyset
                    Do While Not .EOF
                        gcnOracle.Execute CStr(.Fields(0).value)
                        .MoveNext
                    Loop
                End With
                gcnOracle.Execute "truncate table " & mstr所有者 & "." & strTable & " drop storage"
            End If
            
            strCommand = "SQLLDR" & mstrVer & " " & gstrUserName & "/" & gstrPassword & IIf(gstrServer = "", "", "@" & gstrServer) _
                & " CONTROL=" & strPath & strTable & ".LDR DIRECT=TRUE"
            On Error Resume Next
            frmMDIMain.stbThis.Panels(2).Text = strCommand
            lngTemp = Shell(strCommand, vbHide)
            If err = 0 Then
                On Error GoTo 0
                lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
                Do
                    Sleep 100
                    GetExitCodeProcess lngProcess, lngTemp
                    DoEvents
                Loop While lngTemp = Still_Active
                CloseHandle lngProcess
                
                If lngTemp <> 0 And lngTemp <> 1 Then
                    frmWait.EndWait
                    MsgBox "数据调入程序运行失败。", vbCritical, gstrSysName
                    SetEnable True
                    Exit Sub
                End If
                
            Else
                strErrTabs = strErrTabs & vbCrLf & "○" & strTable
            End If
        End If
    Next
    frmMDIMain.stbThis.Panels(2).Text = ""
    '恢复所有失效约束(enable)
    gstrSQL = "select 'ALTER TABLE '||table_name||' ENABLE CONSTRAINT '||constraint_name,constraint_name" & _
            " from user_constraints" & _
            " where STATUS='DISABLED'"
    With rsTemp
        If .State = adStateOpen Then .Close
        .Open gstrSQL, gcnOracle, adOpenKeyset
        On Error Resume Next
        Do Until .EOF
            gcnOracle.Execute CStr(.Fields(0).value)
            If err <> 0 Then
                strErrCons = strErrCons & vbCrLf & "○" & .Fields(1).value
            End If
            .MoveNext
        Loop
    End With
    
    
    '恢复序列
    Call AdjustSequence(mstr所有者, gcnOracle)
    frmWait.EndWait
    SetEnable True
    '总结
    If strErrTabs <> "" Then
        strErrTabs = vbCrLf & "由于文件错误，以下数据表无法执行Loader:" & vbCrLf & strErrTabs
    End If
    If strErrCons <> "" Then
        strErrTabs = strErrTabs & vbCrLf & "由于数据原因，以下约束现在无效，请检查:" & vbCrLf & strErrCons
    End If
    MsgBox "数据调入完毕！" & vbCrLf & vbCrLf & _
        "共耗时" & Format(CDate(Now - varTime), "hh时mm分ss秒。") & _
        IIf(strErrTabs = "", "", "但是" & strErrTabs), vbExclamation, gstrSysName
End Sub

Private Sub SetEnable(ByVal blnEnable As Boolean)
    frmMDIMain.Enabled = blnEnable
    lvwTabs.Enabled = blnEnable
    DirBak.Enabled = blnEnable
    DriveBak.Enabled = blnEnable
    chkDelete.Enabled = blnEnable
    cmdClear.Enabled = blnEnable
    CmdSelect.Enabled = blnEnable
    cmdExecute.Enabled = blnEnable
End Sub

Private Sub Form_Load()
    Dim intVer As Integer
    
    intVer = GetOracleVersion
    
    If intVer < 80 Then
        MsgBox "该Oracle版本可能由于过旧，本程序可能不能正常运行，" & vbCr _
            & "请考虑将BIN目录中的[IMP+版本号.EXE]改为[IMP.EXE]再执行。", vbExclamation, gstrSysName
        mstrVer = ""
    ElseIf intVer = 80 Then            'Oracle8.0
        mstrVer = "80"
    Else
        mstrVer = ""
    End If
    Call FillSystem
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsTable.State = 1 Then mrsTable.Close
    Set mrsTable = Nothing
    mstr所有者 = ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtFile.Width = ScaleWidth - 200 - txtFile.Left
    txtFile.Height = ScaleHeight - 200 - txtFile.Top
    
End Sub

Private Sub FillSystem()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    '显示可显示的系统
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", UCase(gstrUserName))
    
    If Not rsTemp.EOF Then
        cmbSystem.AddItem rsTemp("名称") & " v" & rsTemp("版本号") & "（" & rsTemp("编号") & "）"
        mstr所有者 = UCase(gstrUserName)
        
        mrsTable.CursorLocation = adUseClient
        mrsTable.Open "select table_name from all_tables where owner='" & mstr所有者 & "'", gcnOracle, adOpenStatic, adLockReadOnly
        
        cmbSystem.ListIndex = 0
    Else
        cmbSystem.Enabled = False
        cmdExecute.Enabled = False
        DriveBak.Enabled = False
        DirBak.Enabled = False
        lvwTabs.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub

