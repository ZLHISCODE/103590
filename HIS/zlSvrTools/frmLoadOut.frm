VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoadOut 
   BackColor       =   &H80000005&
   Caption         =   "数据调出"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmLoadOut.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   7320
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExecute 
      Caption         =   "执行(&E)"
      Height          =   350
      Left            =   4035
      TabIndex        =   8
      Top             =   4980
      Width           =   1155
   End
   Begin VB.DirListBox DirBak 
      Appearance      =   0  'Flat
      Height          =   3240
      Left            =   3735
      TabIndex        =   7
      Top             =   1605
      Width           =   3180
   End
   Begin VB.DriveListBox DriveBak 
      Height          =   300
      Left            =   3735
      TabIndex        =   6
      Top             =   1290
      Width           =   3180
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   1065
      TabIndex        =   10
      Top             =   4980
      Width           =   1080
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全清(&R)"
      Height          =   350
      Left            =   2325
      TabIndex        =   11
      Top             =   4980
      Width           =   1080
   End
   Begin VB.CommandButton cmdSQL 
      Caption         =   "生成S&QL脚本…"
      Height          =   375
      Left            =   5325
      TabIndex        =   9
      Top             =   4950
      Width           =   1575
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   690
      Width           =   4815
   End
   Begin MSComctlLib.ListView lvwTabs 
      Height          =   3555
      Left            =   990
      TabIndex        =   4
      Top             =   1290
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   6271
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
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件存放目录(&D)"
      Height          =   180
      Left            =   3750
      TabIndex        =   5
      Top             =   1080
      Width           =   1350
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
   Begin VB.Label lbl说明 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   990
      TabIndex        =   12
      Top             =   5490
      Width           =   6195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "调出系统(&S)"
      Height          =   180
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   750
      Width           =   990
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   240
      Picture         =   "frmLoadOut.frx":04F9
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据调出"
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
Attribute VB_Name = "frmLoadOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mrsSystem As New ADODB.Recordset
Dim mstr所有者 As String '保存当前系统的所有者名

Private Sub cmdSelect_Click()
    Dim i As Integer
    For i = 1 To lvwTabs.ListItems.Count
        lvwTabs.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    For i = 1 To lvwTabs.ListItems.Count
        lvwTabs.ListItems(i).Checked = False
    Next
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
'
'Private Sub chkDate_Click()
'    If chkDate.Value = 1 Then
'        dtpStart.Enabled = True
'        dtpEnd.Enabled = True
'    Else
'        dtpStart.Enabled = False
'        dtpEnd.Enabled = False
'    End If
'End Sub
'
'Private Sub dtpEnd_Change()
'    dtpStart.MaxDate = dtpEnd.Value
'    If dtpStart.Value > dtpStart.MaxDate Then dtpStart.Value = dtpStart.MaxDate
'End Sub
'
'Private Function GetWhere(ByVal strTable As String) As String
'    '时间范围
'    Dim strWhere As String
'
'    strWhere = ""
'    If chkDate.Value = 1 Then
'        Select Case strTable
'            Case "病案主页"
'                strWhere = "入院日期"
'            Case "病人床日费用"
'                strWhere = "开始日期"
'            Case "病人床位费用"
'                strWhere = "开始日期"
'            Case "病人床位记录"
'                strWhere = "入住时间"
'            Case "病人费用记录"
'                strWhere = "登记时间"
'            Case "病人结帐记录"
'                strWhere = "收费时间"
'            Case "病人入出记录"
'                strWhere = "入科时间"
'            Case "病人收治费用"
'                strWhere = "开始日期"
'            Case "病人信息"
'                strWhere = "登记时间"
'            Case "病人预交记录"
'                strWhere = "收款时间"
'            Case "部门表"
'                strWhere = "建档时间"
'            Case "床位等级"
'                strWhere = "建档时间"
'            Case "床位增减"
'                strWhere = "记录日期"
'            Case "给药途径"
'                strWhere = "建档时间"
'            Case "功能内容表"
'                strWhere = "建档日期"
'            Case "挂号项目"
'                strWhere = "建档时间"
'            Case "合约单位"
'                strWhere = "建档时间"
'            Case "护理等级"
'                strWhere = "建档时间"
'            Case "门诊病案记录"
'                strWhere = "建立日期"
'            Case "票据重打记录"
'                strWhere = "打印时间"
'            Case "收费价目"
'                strWhere = "执行日期"
'            Case "收费细目"
'                strWhere = "建档时间"
'            Case "收入项目"
'                strWhere = "建档时间"
'            Case "未发药品记录"
'                strWhere = "填制日期"
'            Case "药品采购计划"
'                strWhere = "编制日期"
'            Case "药品短损记录"
'                strWhere = "登记时间"
'            Case "药品付款记录"
'                strWhere = "填制日期"
'            Case "药品供应商"
'                strWhere = "建档时间"
'            Case "药品目录"
'                strWhere = "建档时间"
'            Case "药品收发记录"
'                strWhere = "审核日期"
'            Case "医嘱记录"
'                strWhere = "登记时间"
'            Case "诊疗项目"
'                strWhere = "建档时间"
'            Case "住院病案记录"
'                strWhere = "建立日期"
'        End Select
'        If strWhere <> "" Then
'            strWhere = " Where " & strWhere & " Between To_Date('" & Format(dtpStart, "yyyy-MM-dd HH:mm:ss") _
'                    & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd, "yyyy-MM-dd HH:mm:ss") _
'                    & "','YYYY-MM-DD HH24:MI:SS')"
'        End If
'    End If
'    GetWhere = strWhere
'End Function

Private Sub SetEnable(ByVal blnEnable As Boolean)
    frmMDIMain.Enabled = blnEnable
    lvwTabs.Enabled = blnEnable
    DirBak.Enabled = blnEnable
    DriveBak.Enabled = blnEnable
    cmdClear.Enabled = blnEnable
    cmdSelect.Enabled = blnEnable
    cmdExecute.Enabled = blnEnable
    cmdSQL.Enabled = blnEnable
End Sub

Private Sub cmdExecute_Click()
    Dim blnGen As Boolean
    Dim strBakDir As String
    Dim strLDR As String
    Dim strField As String
    Dim strTable As String
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    blnGen = False
    If MsgBox("如果用户数据过大，该程序运行速度会比较慢！" & vbCrLf & "要继续执行吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    SetEnable False
    strBakDir = DirBak.Path
    If Right(strBakDir, 1) <> "\" Then strBakDir = strBakDir & "\"
    
    MousePointer = 11
    For Each objItem In lvwTabs.ListItems
        If objItem.Checked Then
            strTable = objItem.Text
            frmMDIMain.stbThis.Panels(2).Text = "调出" & strTable & "："
            gstrSQL = "select COLUMN_NAME,DATA_TYPE " & _
                    " from all_tab_columns " & _
                    " where owner='" & mstr所有者 & "' and table_name='" & strTable & "'" & _
                    "       and DATA_TYPE not in('LONG','LONG RAW','CLOB','BLOB','BFILE','NCLOB','NBLOB')" & _
                    " order by column_id"
            With rsTemp
                If .State = adStateOpen Then .Close
                .Open gstrSQL, gcnOracle, adOpenKeyset
                
                strLDR = ""
                strField = ""
                Do Until .EOF
                    If .Fields(1).value = "DATE" Then
                        strLDR = strLDR & "||'^'||To_Char(""" & .Fields(0).value & """,'YYYY-MM-DD HH24:MI:SS')"
                        strField = strField & ",""" & .Fields(0).value & """ Date 'YYYY-MM-DD HH24:MI:SS' "
                    Else
                        strLDR = strLDR & "||'^'||""" & .Fields(0).value & """"
                        strField = strField & ",""" & .Fields(0).value & """"
                    End If
                    .MoveNext
                Loop
                
                '查询SQL语句
                If strLDR = "" Then
                    objItem.Checked = False
                Else
                    blnGen = True
                    gstrSQL = "select " & Mid(strLDR, 8) & " from " & mstr所有者 & "." & strTable ' & GetWhere(strTable)
                    If .State = adStateOpen Then .Close
                    .Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
                    If .RecordCount <> 0 Then
                        On Error Resume Next
                        Kill strBakDir & strTable & ".ldr"
                        On Error GoTo 0
                        Open strBakDir & strTable & ".ldr" For Binary Access Write As #1
                        strLDR = "LOAD DATA INFILE *" & vbCrLf & _
                                "PRESERVE BLANKS" & vbCrLf & _
                                "INTO TABLE " & strTable & " APPEND" & vbCrLf & _
                                "FIELDS TERMINATED BY '^'" & vbCrLf & _
                                "TRAILING NULLCOLS(" & Mid(strField, 2) & ")" & vbCrLf & _
                                "BEGINDATA"
                        Put #1, , strLDR & vbCrLf
                        Do Until .EOF
                            Put #1, , Replace(CStr(.Fields(0).value), vbCrLf, vbCr) & vbCrLf
                            If Int(.AbsolutePosition / .RecordCount * 1000) Mod 10 = 0 Then
                                frmMDIMain.stbThis.Panels(2).Text = "调出" & strTable & "：" & String(Int(.AbsolutePosition * 16 / .RecordCount), "…")
                                DoEvents
                            End If
                            .MoveNext
                        Loop
                        Close #1
                    Else
                        objItem.Checked = False
                    End If
                End If
            End With
        End If
    Next
    MousePointer = 0
    frmMDIMain.stbThis.Panels(2).Text = ""
    SetEnable True
        
    If Not blnGen Then
        MsgBox "由于没有选择表或选择的表没有标量字段，无法生成Loader文件。", vbExclamation, gstrSysName
    Else
        MsgBox "Loader文件生成完毕。" & vbCr _
            & vbCr & "如果发现选中的数据表不再是选中状态" _
            & vbCr & "说明该表没有数据或没有简单数据类型列！", vbExclamation, gstrSysName
    End If
End Sub

Private Sub cmdSQL_Click()
    Dim blnGen As Boolean
    Dim strBakDir As String
    Dim strLDR As String
    Dim strField As String
    Dim strTable As String
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    blnGen = False
    
    SetEnable False
    strBakDir = DirBak.Path
    If Right(strBakDir, 1) <> "\" Then strBakDir = strBakDir & "\"
    
    '创建CreateLoaderData.Sql文件
    On Error Resume Next
    Kill strBakDir & "CreateLoaderData.Sql"
    
    '添加表头内容
    On Error GoTo 0
    Open strBakDir & "CreateLoaderData.Sql" For Binary Access Write As #1
    Put #1, , "set echo off heading off feedback off verify off;" & vbCrLf
    Put #1, , "set linesize 30000 pagesize 0 trimspool on;" & vbCrLf
    Put #1, , "set termout off;" & vbCrLf & vbCrLf
    
    For Each objItem In lvwTabs.ListItems
        If objItem.Checked Then
            strTable = objItem.Text
            gstrSQL = "select COLUMN_NAME,DATA_TYPE " & _
                    " from all_tab_columns " & _
                    " where owner='" & mstr所有者 & "' and table_name='" & strTable & "'" & _
                    "       and DATA_TYPE not in('LONG','LONG RAW','CLOB','BLOB','BFILE','NCLOB','NBLOB')" & _
                    " order by column_id"
            
            With rsTemp
                If .State = adStateOpen Then .Close
                .Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
                
                strLDR = ""
                strField = ""
                Do While Not .EOF
                    If .Fields(1).value = "DATE" Then
                        strLDR = strLDR & "||'^'||To_Char(""" & .Fields(0).value & """,'YYYY-MM-DD HH24:MI:SS')"
                        strField = strField & ",""" & .Fields(0).value & """ Date 'YYYY-MM-DD HH24:MI:SS' "
                    Else
                        strLDR = strLDR & "||'^'||""" & .Fields(0).value & """"
                        strField = strField & ",""" & .Fields(0).value & """"
                    End If
                    .MoveNext
                Loop
                If strLDR = "" Then
                    objItem.Checked = False
                Else
                    blnGen = True
                    gstrSQL = "select " & Mid(strLDR, 8) & " from " & mstr所有者 & "." & strTable
                    Put #1, , "spool " & strBakDir & strTable & ".ldr;" & vbCrLf
                    '控制文件表头开始
                    Put #1, , "select 'LOAD DATA INFILE *' from dual;" & vbCrLf
                    Put #1, , "select 'PRESERVE BLANKS' from dual;" & vbCrLf
                    Put #1, , CStr("select 'INTO TABLE " & strTable & " APPEND' from dual;") & vbCrLf
                    Put #1, , "select 'FIELDS TERMINATED BY ''^''' from dual;" & vbCrLf
                    Put #1, , CStr("select 'TRAILING NULLCOLS(" _
                        & Replace(Mid(strField, 2), "'", "''") & ")' from dual;") & vbCrLf
                    Put #1, , "select 'BEGINDATA' from dual;" & vbCrLf
                    '控制文件表头结束
                    
                    '查询SQL语句
                    Put #1, , gstrSQL & ";" & vbCrLf ' & GetWhere(strTable)
                    Put #1, , "spool off;" & vbCrLf & vbCrLf
                End If
            End With
            frmMDIMain.stbThis.Panels(2) = "正在生成《" & strTable & "》表的脚本……"
        End If
    Next
    Put #1, , "exit" & vbCrLf
    Close #1
    frmMDIMain.stbThis.Panels(2) = ""
    SetEnable True
    
    If Not blnGen Then
        MsgBox "由于没有选择表或选择的表没有标量字段，无法生成脚本文件。", vbExclamation, gstrSysName
    Else
        If MsgBox("SQL脚本已经生成完毕。请查阅“" & strBakDir & "CreateLoaderData.Sql”文件。" & vbCrLf & vbCrLf & _
               "你想马上就运行该脚本吗？", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
            
            Dim lngTemp As Long
            Dim lngProcess  As Long
            
            frmWait.BeginWait "正在生成SQL脚本……"
            lngTemp = Shell("sqlplus " & gstrUserName & "/" & gstrPassword & IIf(gstrServer = "", "", "@" & gstrServer) & " @" & strBakDir & "CreateLoaderData.sql", vbHide)
            If err <> 0 Then
                err.Clear
                MsgBox "不能正确生成脚本，请检查：" & _
                    vbCr & "   1) 是否存在sqlplus.exe文件；" & _
                    vbCr & "   2) Set Path是否指向其存在的目录。", vbExclamation, gstrSysName
                frmWait.EndWait
                Exit Sub
            End If
            
            lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
            Do
                Sleep 100
                GetExitCodeProcess lngProcess, lngTemp
                DoEvents
            Loop While lngTemp = Still_Active
            CloseHandle lngProcess
            frmWait.EndWait
                
            If lngTemp <> 0 Then
                MsgBox "脚本生成程序非法退出。", vbCritical, gstrSysName
            End If
        End If
    End If
End Sub

Private Sub cmbSystem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then lvwTabs.SetFocus
End Sub

Private Sub Form_Load()
    lbl说明.Caption = "说明：" & vbCrLf & _
                    "     调出数据要经过一个很漫长的过程才能完成。在这段时间内服务器对客户的响应会变得迟钝，因最好在服务器空闲时完成本操作。" & vbCrLf & _
                    "     对于每个数据表都会产生一个同名的导出文件，导入时可根据这些文件独立完成。" & vbCrLf & _
                    "     用SQLPLUS运行生成的脚本来得到导出文件，可能比本程序执行的效率更高，因此建议使用脚本方式。"
'    dtpStart.Value = Format(Date & " " & Format("0:0:0", "HH:mm:ss"), "yyyy-MM-dd HH:mm:ss")
'    dtpEnd.Value = Format(Date & " " & Format("23:59:59", "HH:mm:ss"), "yyyy-MM-dd HH:mm:ss")
    Call FillSystem
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsSystem.State = 1 Then mrsSystem.Close
    Set mrsSystem = Nothing
    mstr所有者 = ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lbl说明.Width = ScaleWidth - 200 - lbl说明.Left
    lbl说明.Height = ScaleHeight - 200 - lbl说明.Top
    
End Sub

Private Sub cmbSystem_Click()
    Dim rsTemp As New ADODB.Recordset
    
    mrsSystem.Filter = "编号=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    lvwTabs.ListItems.Clear
    If mrsSystem.RecordCount = 0 Then
        cmdExecute.Enabled = False
    Else
        cmdExecute.Enabled = True
        mstr所有者 = mrsSystem("所有者")
        
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open "select table_name from all_tables where owner='" & mstr所有者 & "'", gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            lvwTabs.ListItems.Add , , rsTemp("TABLE_NAME")
            rsTemp.MoveNext
        Loop
    End If
End Sub

Private Sub FillSystem()
    '显示所有可显示的系统
    On Error GoTo errHandle
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

