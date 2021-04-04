VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppCheck 
   BackColor       =   &H80000005&
   Caption         =   "对象检查修复"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmAppCheck.frx":0000
   ScaleHeight     =   6135
   ScaleWidth      =   5610
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFunction 
      Caption         =   "对象权限修正(&O)"
      Height          =   350
      Index           =   6
      Left            =   825
      TabIndex        =   17
      Top             =   3555
      Width           =   1650
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "历史库修正(&H)"
      Height          =   350
      Index           =   5
      Left            =   825
      TabIndex        =   16
      Top             =   3550
      Width           =   1650
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "同义词修正(&N)"
      Height          =   350
      Index           =   4
      Left            =   825
      TabIndex        =   15
      Top             =   3210
      Width           =   1650
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "整理序列(&S)"
      Height          =   350
      Index           =   3
      Left            =   825
      TabIndex        =   10
      Top             =   2880
      Width           =   1650
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   5550
      TabIndex        =   12
      Top             =   5595
      Visible         =   0   'False
      Width           =   5610
      Begin MSComctlLib.ProgressBar pgbState 
         Height          =   180
         Left            =   135
         TabIndex        =   13
         Top             =   255
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正在检查"
         Height          =   180
         Left            =   135
         TabIndex        =   14
         Top             =   60
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "重建索引(&I)"
      Height          =   350
      Index           =   2
      Left            =   825
      TabIndex        =   9
      Top             =   2550
      Width           =   1650
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "对象修正(&R)"
      Height          =   350
      Index           =   1
      Left            =   825
      TabIndex        =   3
      Top             =   2220
      Width           =   1650
   End
   Begin VB.CommandButton cmdGetIni 
      Caption         =   "选择(&S)…"
      Height          =   350
      Left            =   4215
      TabIndex        =   5
      Top             =   1020
      Width           =   1095
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   645
      Width           =   3570
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "对象检查(&C)"
      Height          =   350
      Index           =   0
      Left            =   825
      TabIndex        =   2
      Top             =   1860
      Width           =   1650
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "将现有系统的数据库对象与安装文件对比，检查现有系统的序列、表、视图、索引、存储过程等对象的正确性。"
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   2565
      TabIndex        =   11
      Top             =   1905
      Width           =   2730
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   840
      Left            =   900
      TabIndex        =   8
      Top             =   4170
      Width           =   4275
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   960
      TabIndex        =   7
      Top             =   1380
      Width           =   4350
   End
   Begin VB.Label lblFileCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "安装配置文件"
      Height          =   180
      Left            =   960
      TabIndex        =   6
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
      Caption         =   "对象检查修复"
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
      Picture         =   "frmAppCheck.frx":04F9
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmAppCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrIniPath    As String                 '安装配置文件目录
Private mrsErrTable As ADODB.Recordset
Private mrsExeErrTable As ADODB.Recordset

Private WithEvents mclsObjectCheck As clsObjectCheck
Attribute mclsObjectCheck.VB_VarHelpID = -1
Private mclsRunScript As New clsRunScript

Private Enum CMDFUN
    E对象检查 = 0
    E对象修正 = 1
    E重建索引 = 2
    E重整序列 = 3
    E同义词 = 4
    E历史结构 = 5
    E对象权限 = 6
End Enum

Private Sub cmbSystem_Click()
    Dim strFilePath As String
    Dim blnTools As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If Val(cmbSystem.ItemData(cmbSystem.ListIndex)) = -1 Then
        strFilePath = App.Path & "\Tools\zlServer.Sql"
        cmbSystem.Tag = "ZLTOOLS"
    Else
        cmbSystem.Tag = GetOwnerName(Val(cmbSystem.ItemData(cmbSystem.ListIndex)), gcnOracle)
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Loadandunload.Get_Sysfile_name", Val(cmbSystem.ItemData(cmbSystem.ListIndex)), 1)
        If Not rsTemp.EOF Then strFilePath = rsTemp.Fields(0).value & ""
    End If
    '设置功能状态
    Call SetFunsState(strFilePath, False)
    
End Sub

Private Function zl获取关联外键(ByVal str对象名 As String, ByRef rsData As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定主键的关联外键
    '入参:str对象名-对象名称
    '出参:rsData
    '返回:
    '编制:刘兴洪
    '日期:2009-08-20 11:30:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset, strTemp As String


    If gblnDBA Then
        strSQL = "Select table_Name,Constraint_Name,OWNER,R_OWNER,R_Constraint_Name,DELETE_RULE from DBA_CONSTRAINTS Where R_Constraint_Name='" & str对象名 & "' And Constraint_Type='R'"
    Else
        strSQL = "Select table_Name,Constraint_Name,OWNER,R_OWNER,R_Constraint_Name,DELETE_RULE From USER_CONSTRAINTS Where  R_Constraint_Name='" & str对象名 & "'  And Constraint_Type='R'"
    End If

    With rsTemp
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        If .RecordCount = 0 Then zl获取关联外键 = True: Exit Function
        .Filter = "OWNER<>'" & cmbSystem.Tag & "' AND R_OWNER='" & cmbSystem.Tag & "'"
        If .RecordCount <> 0 Then
            '存在其他系统关联,当项主键只能手工更新
            If Not rsData.EOF Then
                If UCase(Nvl(rsData!对象名称)) <> UCase(str对象名) Then
                    rsData.Filter = "对象名称='" & UCase(str对象名) & "'"
                End If
            Else
                rsData.Filter = "对象名称='" & UCase(str对象名) & "'"
            End If
            If rsData.EOF Then rsData.Filter = 0: zl获取关联外键 = True: Exit Function
            .MoveFirst: strTemp = ""
            Do While Not .EOF
                strTemp = strTemp & "," & Nvl(!Table_Name) & "(" & Nvl(!Owner) & ")"
                .MoveNext
            Loop
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            If strTemp <> "" Then strTemp = Substr("被其他系统关联，不能有效修正！如表:" & strTemp, 1, 500)
            '更新标志
            rsData!修正标志 = 4
            rsData!修正说明 = strTemp
            rsData.Update
            .Close
            zl获取关联外键 = True: Exit Function
        End If
        .Filter = 0: .MoveFirst

        '需要删除相关的级联
        Do While Not .EOF
            '标志: '1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致,8-处理主键时，需要先处理外键
            Call zlInsertRecData(rsData, Nvl(!Table_Name), Nvl(!Constraint_Name), "外键", 8, False, "", "级联处理", "在处理主键或唯一键时，需要先处理外键！")
            .MoveNext
        Loop
        .Close
        zl获取关联外键 = True: Exit Function
    End With
End Function

Private Sub ModifyToolsObject(ByVal cnTools As ADODB.Connection, ByVal blnDele As Boolean, ByVal strSeverScrip As String)
    '-----------------------------------------------------------------------------------------------------------
    '功能:修正ZLTOOLS相关对象
    '参数:cnTools-连接到管理工具的连接
    '     blnDele-是否删除相关的约束等
    '     strSeverScrip-脚务器管理工具脚本文件
    '返回:
    '编制:刘兴宏
    '日期:2007/09/12
    '-----------------------------------------------------------------------------------------------------------
        Dim rsTemp As New ADODB.Recordset
    Dim rsSys As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    '根据安装文件尽量修正
    '一、序列：直接重新建立，出错继续
    '三、约束：删除现有约束重新建立
    '四、索引：删除现有索引重新建立
    '五、视图：直接重新建立
    '六、程序：直接重新建立
    '七、同义词：根据工具对象，修正同义词
    
    If blnDele Then
        lblStatus.Caption = "删除外键与检查约束…"
        strSQL = "select 'alter table '||table_name||' drop constraint '||constraint_name" & _
                " From user_constraints" & _
                " where constraint_type IN ('R','C') And Instr(Table_Name,'BIN$')<=0 And Instr(constraint_name,'SYS_')<=0 "
        OpenRecordset rsTemp, strSQL, Me.Caption, , , cnTools
        With rsTemp
            Do While Not .EOF
                cnTools.Execute .Fields(0).value
                pgbState.value = .AbsolutePosition / .RecordCount * 100
                DoEvents
                .MoveNext
            Loop
        End With
        lblStatus.Caption = "删除主键与唯一约束…"
        strSQL = "select 'alter table '||table_name||' drop constraint '||constraint_name" & _
                " From user_constraints Where Instr(Table_Name,'BIN$')<=0 And Instr(constraint_name,'SYS_')<=0 "
        OpenRecordset rsTemp, strSQL, Me.Caption, , , cnTools
        With rsTemp
            Do While Not .EOF
                cnTools.Execute .Fields(0).value
                pgbState.value = .AbsolutePosition / .RecordCount * 100
                DoEvents
                .MoveNext
            Loop
        End With
        lblStatus.Caption = "删除索引…"
        strSQL = "select 'drop Index ""'||Index_name||'""'  from user_indexes Where INDEX_TYPE='NORMAL' And Instr(Table_Name,'BIN$')<=0"
        OpenRecordset rsTemp, strSQL, Me.Caption, , , cnTools
        With rsTemp
            Do While Not .EOF
                cnTools.Execute .Fields(0).value
                pgbState.value = .AbsolutePosition / .RecordCount * 100
                DoEvents
                .MoveNext
            Loop
        End With
    End If
    '根据安装脚本重新执行
    
    err = 0: On Error Resume Next
    
    If Not mclsRunScript.OpenFile(strSeverScrip) Then Exit Sub
    lblStatus.Caption = "正在修正对象…"
    Do While Not mclsRunScript.EOF
        strSQL = mclsRunScript.SQLInfo.SQL
'        If mclsRunScript.Line >= 25239 Then Stop
        If Not mclsRunScript.SQLInfo.Block Then
            strTmp = mclsRunScript.SQLInfo.PartSQL
            If strTmp Like "CREATE TABLE *" Or strTmp Like "CREATE GLOBAL TEMPORARY TABLE *" Then
              '不处理
            ElseIf strTmp Like "ALTER TABLE * CONSTRAINT *" Then
                '修正约束
                    cnTools.Execute strSQL
            ElseIf strTmp Like "CREATE INDEX *" Then
                '修正索引
                cnTools.Execute strSQL
            ElseIf strTmp Like "CREATE SEQUENCE *" Then
                '检查序列
                cnTools.Execute strSQL
            End If
        Else
            strTmp = mclsRunScript.SQLInfo.BlockType
            If strTmp = "TYPE" Then
            
                '不处理
            ElseIf strTmp Like "*PROCEDURE*" Or _
                    strTmp Like "*FUNCTION*" Or _
                    strTmp Like "*PACKAGE*" Then
                '检查过程与函数的合法性
                cnTools.Execute strSQL
            End If
        End If
        err.Clear
        pgbState.value = mclsRunScript.ProcessValue
        Call mclsRunScript.ReadNextSQL
        DoEvents
    Loop
    lblStatus.Caption = "正在修正共公同义词…"
    Call ReGrantForTools(cnTools, , True)
    lblStatus.Caption = "正在重整序列…"
    Call AdjustSequence("ZLTOOLS", cnTools)
End Sub
Private Function zlCheckTableDataIsNull(ByVal strTableName As String, ByVal str字段名 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定数据表的字定字段是否为NULL
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-08-20 14:07:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0: On Error Resume Next
    
    strSQL = "Select 1 From " & cmbSystem.Tag & "." & strTableName & " where " & str字段名 & " is not null and rownum<=1"
    rsTemp.Open strSQL, gcnOracle
    zlCheckTableDataIsNull = Not (rsTemp.RecordCount <> 0)
    
End Function

Private Function zlModifyObject(ByVal str类型 As String, Optional blnDelete As Boolean = True) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:数据对象修正
        '入参:str类型-指定的类型:外键;约束;索引
        '     blnDelete-是否删除
        '出参:
        '返回:
        '编制:刘兴洪
        '日期:2009-08-19 17:35:10
        '---------------------------------------------------------------------------------------------------------------------------------------------
         Dim i As Long, strSQL As String, str修改说明 As String, byt修正标志 As Integer, varData As Variant, lng长度 As Long, lng精度 As Long
         Dim iCount As Integer
         
        If blnDelete Then
            lblStatus.Caption = "正在删除错误" & str类型 & "…"
        Else
            lblStatus.Caption = "正在修正错误" & str类型 & "…"
        End If
        mrsErrTable.Filter = "类型='" & str类型 & "'" & IIf(blnDelete, "", " and 修正标志<>2")
        With mrsErrTable
            i = 0: iCount = .RecordCount
            Do While Not .EOF
            
                strSQL = Nvl(!修正语句)
                Select Case str类型
                Case "约束", "外键"
                   If blnDelete Then strSQL = " Alter table " & Nvl(!表名称) & "  Drop constraint " & Nvl(!对象名称)
                Case "索引"
                    If blnDelete Then strSQL = " Drop index " & Nvl(!对象名称)
                Case "数据表"
                    '需要检查,是否修正标志为手工执行的
                     If !修正标志 = 4 And blnDelete = False And Nvl(!原字段类型) <> "" Then  '0-未修正,1-已经修正,2-修正失败,4-不能执行修正，需要手工调整
                        '精度小的话,需要检查数据是否为空,不为空,则可以升级
                        varData = Split(Nvl(!原字段长度) & ",", ",")
                        lng长度 = Val(varData(0)): lng精度 = Val(varData(1))
                        varData = Split(Nvl(!现字段长度) & ",", ",")
                        If lng长度 > Val(varData(0)) Or lng精度 > Val(varData(1)) Then
                           If zlCheckTableDataIsNull(Nvl(!表名称), Nvl(!字段名)) Then
                                str修改说明 = "虽然精度改小,但由于字段内容为空,因此,也修正了该结构!"
                           Else '不为空,不能更改
                                strSQL = ""
                           End If
                        End If
                     End If
                End Select
                str修改说明 = "修正成功": byt修正标志 = 1
                
                If strSQL <> "" Then
                    err = 0: On Error Resume Next
                    gcnOracle.Execute strSQL
                    If err <> 0 Then
                          byt修正标志 = 2
                         Select Case str类型
                         Case "约束"
                            str修改说明 = Substr("不能有效修正主键或唯一键,错误信息:" & err.Description, 1, 4000)
                         Case "外键"
                            str修改说明 = Substr("不能有效修正外键,错误信息:" & err.Description, 1, 4000)
                         Case "索引"
                            str修改说明 = Substr("不能有效修正索引,错误信息:" & err.Description, 1, 40000)
                         Case "数据表"
                            str修改说明 = Substr("不能有效修正数据表,错误信息:" & err.Description, 1, 40000)
                         End Select
                    End If
                    
                    If blnDelete = False Then
                        !修正说明 = str修改说明
                        !修正标志 = byt修正标志
                        .Update
                    End If
                    err = 0: On Error GoTo 0
                End If
                i = i + 1
                pgbState.value = i / iCount * 100
                DoEvents
                .MoveNext
            Loop
            .Filter = 0
        End With
End Function
Private Sub cmdFunction_Click(Index As Integer)
    Dim i As Long, intVer As Integer
    Dim bytOperation As VbMsgBoxResult
    Dim blnTool As Boolean
    Dim cnTools As ADODB.Connection
    Dim lngSys As Long, lngAbort As Long
    Dim cllSQL As New Collection, cllErr As New Collection
    Dim strTmp As String
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    Dim rsObjects As New ADODB.Recordset
    
    Call cmdFunction_MouseMove(Index, 0, 0, 0, 0)
    
    If MsgBox("""" & Split(cmdFunction(Index).Caption, "(")(0) & """操作将可能消耗较多的资源和花费较长的时间，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    intVer = GetOracleVersion
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 0

        Set mclsObjectCheck = New clsObjectCheck
   
        If cmbSystem.Tag = "ZLTOOLS" Then
            
            If mclsObjectCheck.InitCheckManageTool(Me, lblFileName.Caption) = True Then
                Call mclsObjectCheck.CheckToolsObject
                Call mclsObjectCheck.ShowReport
            End If

        Else
            
            If mclsObjectCheck.InitCheck(Me, cmbSystem.Tag, lblFileName.Caption, gstrUserName, gstrSysName, gblnDBA, cmbSystem.ItemData(cmbSystem.ListIndex), cmbSystem.Text) = True Then
                Call mclsObjectCheck.CheckObject
                Call mclsObjectCheck.ShowReport
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        blnTool = (cmbSystem.Tag = "ZLTOOLS")
        If blnTool Then
            If gobjFile.FileExists(lblFileName.Caption) = False Then
                For i = 0 To cmdFunction.Ubound
                    cmdFunction(i).Enabled = False
                Next
                lblFileName.Caption = ""
                cmdGetIni.SetFocus
                Exit Sub
            End If
            
            Set cnTools = GetConnection("ZLTOOLS")
            If cnTools Is Nothing Then
                MsgBox "打开管理工具失败,请检查!", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            If CheckIniFile(lblFileName.Caption, True) = False Then
                For i = 0 To cmdFunction.Ubound
                    cmdFunction(i).Enabled = False
                Next
                lblFileName.Caption = ""
                cmdGetIni.SetFocus
                Exit Sub
            End If
        End If
        
        '根据安装文件尽量修正
        '一、序列：直接重新建立，出错继续
        '二、数据表：先检查出数据表的错误,对错误的表进行修正(字段不存在,修正;字段类型不对,提醒手工调整;字段精度比脚本的大时,提醒手工调整(但无数据时,自动调整);字段精度比脚本的小,则自动调整)
        '三、约束：删除检查出来的约束,然后重新执行对应的约束脚本
        '四、索引：删除检查出来的索引,然后重新执行对应的索引脚本
        '五、视图：直接重新建立
        '六、程序：直接重新建立
        strTmp = "警告:" & vbCrLf & _
                "    执行对象修正功能,应慎重，为了避免数据丢失，建义在执行该功能前，请作如下的检查：" & vbCrLf & _
                "1. 确认所有的联机用户处于断开状态。" & vbCrLf & _
                "2. 数据应该备份，以便数据恢复。" & vbCrLf & _
                "你是否真的要执行“对象修正”功能？"
        
        bytOperation = MsgBox(strTmp, vbQuestion + vbYesNo + vbDefaultButton3, gstrSysName)
        If bytOperation = vbNo Then Exit Sub
        
        
        Set cllSQL = New Collection
        picStatus.Visible = True
        Enabled = False
        If blnTool Then
            '初始化内部数据集
            Call zlInitRec(mrsErrTable)
            Call ModifyToolsObject(cnTools, bytOperation = vbYes, lblFileName.Caption)
            DoEvents
            Call ReCompileProcedure(cnTools)
        Else
            '初始化内部数据集
            Call zlInitRec(mrsErrTable)
            
            '先检查错误
            lblStatus.Caption = "正在检查数据表…"
            Call CheckTable(mstrIniPath & "zlTable.sql", True)
            lblStatus.Caption = "正在检查约束…"
            Call CheckConstraint(mstrIniPath & "zlConstraint.sql", True)
            
            lblStatus.Caption = "正在检查索引…"
            Call CheckIndex(mstrIniPath & "zlIndex.sql", True)
            
            '外键删除
            lblStatus.Caption = "正在删除外键…"
            Call zlModifyObject("外键", True)
            '约束删除
            lblStatus.Caption = "正在删除约束…"
            Call zlModifyObject("约束", True)
            '索引删除
            lblStatus.Caption = "正在删除索引…"
            Call zlModifyObject("索引", True)
            '修正数据表:
            lblStatus.Caption = "正在修正数据表…"
            Call zlModifyObject("数据表", False)
            '约束修正:
            lblStatus.Caption = "正在修正约束…"
            Call zlModifyObject("约束", False)
            '外键修正:
            lblStatus.Caption = "正在修正外键…"
            Call zlModifyObject("外键", False)
            '索引修正:
            lblStatus.Caption = "正在修正索引…"
            Call zlModifyObject("索引", False)
            
            lngSys = cmbSystem.ItemData(cmbSystem.ListIndex)
            '重新实例化，清除使用痕迹
            Set mclsRunScript = New clsRunScript
            '设置参数类参数
            Call mclsRunScript.InitGlobalPara(Me, lngSys)
            '初始化用户密码信息，加密快可能用到
            Call mclsRunScript.InitUserList(gstrUserName, gstrPassword)
        
            '根据安装脚本重新建立
            lblStatus.Caption = "正在修正序列…"
            Call RunSQLScript(gcnOldOra, mstrIniPath & "zlSequence.sql", True)
            lblStatus.Caption = "正在修正视图…"
            Call RunSQLScript(gcnOldOra, mstrIniPath & "zlView.sql", True)
            lblStatus.Caption = "正在修正函数与过程…"
            Call RunSQLScript(gcnOldOra, mstrIniPath & "zlProgram.sql", True)
            If Dir(mstrIniPath & "zlPackage.sql") <> "" Then
                lblStatus.Caption = "正在修正包…"
                Call RunSQLScript(gcnOldOra, mstrIniPath & "zlPackage.sql", True)
            End If
            lblStatus.Caption = "正在重整序列…"
            Call AdjustSequence(gstrUserName, gcnOldOra, lngSys)
            DoEvents
            Call ReCompileProcedure(gcnOldOra)
        End If
        
        picStatus.Visible = False
        '加载错误信息:刘兴宏
        '出错,给出错信息赋值,以便显示
        
        '"修正标志", adLongVarChar, 2, adFldIsNullable  '0-未修正,1-已经修正,2-修正错败,4-不能执行修正，需要手工调整
        
        If Not mrsErrTable Is Nothing Then
           frmAppChkRpt.blnModiyfyCheck = True
           mrsErrTable.Filter = "修正标志=2"
           With mrsErrTable
               '数据修正错误
               If mrsErrTable.RecordCount <> 0 Then .MoveFirst
               Do While Not .EOF
                   Call InputErrModifyRpt("对象修正错误", Nvl(!类型), Nvl(!对象名称), Nvl(!错误类型), Nvl(!错误信息), "修正错误", Nvl(!修正说明))
                   .MoveNext
               Loop
           End With
           mrsErrTable.Filter = "修正标志=0"
           With mrsErrTable
               '数据修正错误
               If mrsErrTable.RecordCount <> 0 Then .MoveFirst
               Do While Not .EOF
                   Call InputErrModifyRpt("未做调整的对象", Nvl(!类型), Nvl(!对象名称), Nvl(!错误类型), Nvl(!错误信息), "未修正", "")
                   .MoveNext
               Loop
           End With
           
           mrsErrTable.Filter = "修正标志=4"
           With mrsErrTable
               '数据修正错误
               If mrsErrTable.RecordCount <> 0 Then .MoveFirst
               Do While Not .EOF
                   Call InputErrModifyRpt("需要手工修正的对象", Nvl(!类型), Nvl(!对象名称), Nvl(!错误类型), Nvl(!错误信息), "手工调整", Nvl(!修正说明))
                   .MoveNext
               Loop
           End With
           
           mrsErrTable.Filter = "修正标志=1"
           With mrsErrTable
               '数据修正错误
               If mrsErrTable.RecordCount <> 0 Then .MoveFirst
               Do While Not .EOF
                   Call InputErrModifyRpt("修正成功的对象", Nvl(!类型), Nvl(!对象名称), Nvl(!错误类型), Nvl(!错误信息), "修正成功", Nvl(!修正说明))
                   .MoveNext
               Loop
           End With
    
           If frmAppChkRpt.hgdReport.Rows > 1 Then
               If MsgBox("已经执行完全相关对象的修正," & vbCr & "查看检查报告吗？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                   frmAppChkRpt.hgdReport.FixedRows = 1
                   frmAppChkRpt.Show 1
               End If
           Else
               MsgBox "对象修正完毕，可再次运行“对象检查”，确定是否存在未修正对象！", vbInformation, gstrSysName
           End If
           frmAppChkRpt.blnModiyfyCheck = False
        End If
        Enabled = True
    Case 2
        picStatus.Visible = True
        Enabled = False
        lblStatus.Caption = "重建索引…"
        blnTool = (cmbSystem.Tag = "ZLTOOLS")
        With rsTemp
            If gblnDBA Then
                strSQL = "select INDEX_NAME,STATUS from DBA_INDEXES where TABLESPACE_NAME is Not Null And Index_Type = 'NORMAL' And OWNER='" & cmbSystem.Tag & "' and TABLE_OWNER='" & cmbSystem.Tag & "'"
            Else
                strSQL = "select INDEX_NAME,STATUS from USER_INDEXES where TABLESPACE_NAME is Not Null And Index_Type = 'NORMAL' And TABLE_OWNER='" & cmbSystem.Tag & "'"
            End If
            If .State = adStateOpen Then .Close
            .Open strSQL, gcnOracle, adOpenKeyset
            Do While Not .EOF
                lblStatus.Caption = "重建索引：" & .Fields(0).value & "…"
                strSQL = "ALTER INDEX " & cmbSystem.Tag & "." & .Fields(0).value & " Rebuild nologging"
                gcnOldOra.Execute strSQL
                pgbState.value = .AbsolutePosition / .RecordCount * 100
                DoEvents
                .MoveNext
            Loop
        End With
        pgbState.value = 0
        picStatus.Visible = False
        Enabled = True
        MsgBox "索引重建完毕！", vbInformation, gstrSysName
    Case 3
        picStatus.Visible = True
        Enabled = False
        lblStatus.Caption = "正在读取序列…"
        
        Dim lngRealId As Long
        Dim lngNextId As Long
        
        Set rsObjects = GetSequence("", gcnOracle)
        With rsObjects
            Do Until .EOF
                lblStatus.Caption = "整理序列：" & !Sequence_Name & "…"
                pgbState.value = .AbsolutePosition / .RecordCount * 100: DoEvents
                Call AdjustNameSequece(rsObjects!Owner & "." & rsObjects!Table_Name, gcnOracle, rsObjects!Column_Name)
                .MoveNext
            Loop
            
            Call Adjust结帐ID(gcnOracle)
        End With
        pgbState.value = 0
        picStatus.Visible = False
        Enabled = True
        
        MsgBox "序列整理完毕！", vbInformation, gstrSysName
    '修正公共同义词
    Case 4
        '创建当前所有者的全部对象的公共同义词('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
        gcnOracle.Execute "Zl_Createpubsynonyms", , adCmdStoredProc
        
        MsgBox "修正公共同义词完成！", vbInformation, gstrSysName
    Case 5
        Call frmHistorySpaceRepair.ShowRepair(Me, Val(cmbSystem.ItemData(cmbSystem.ListIndex)), False, , , True)
    Case 6
        '对象权限修正
        Set cnTools = GetConnection("ZLTOOLS")
        If cnTools Is Nothing Then Exit Sub
        Call ReGrantForTools(cnTools, , True)
        MsgBox "对象权限修正完成！", vbInformation, gstrSysName
    End Select
End Sub



Private Sub cmdFunction_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngTemp As Long
    Dim i As Long
    
    For i = 0 To cmdFunction.Ubound
        If i = Index Then
            If cmdFunction(Index) Is ActiveControl And cmdFunction(Index).FontBold = True Then Exit Sub
            
            For lngTemp = 0 To cmdFunction.Ubound
                cmdFunction(lngTemp).FontBold = False
            Next
            cmdFunction(i).FontBold = True
            cmdFunction(i).SetFocus
            Select Case i
            Case 0
                lblNote.Caption = "    将现有系统的数据库对象与安装文件对比，检查现有系统的序列、表、视图、索引、存储过程、包等对象的正确性。"
            Case 1
                lblNote.Caption = "    根据安装文件自动修正或重新建立序列、数据表、视图、索引、存储过程、包等数据对象；" & vbCrLf & _
                                  "    为避免修正导致数据不完整，不会对字段类型不一致或从高精度向低精度改变等内容进行修正，同时也就不能保证有效修正所有对象。" & vbCrLf & _
                                  "    该功能仅系统所有者可运行。"
            Case 2
                lblNote.Caption = "    对系统现有的所有索引(包括因主键约束、唯一约束而建立的索引)逐个进行重建操作(Rebuild)，以保证索引的有效性"
            Case 3
                lblNote.Caption = "    调整各个序列的当前值，保证序列与实际应用的匹配；" & vbCrLf & "    当系统出现“…(ID)出现重复！”一类错误时，一般可使用本操作解决问题。"
            Case 4
                lblNote.Caption = "    根据当前所有者的表，视图，过程，函数，序列等对象创建同名公共同义词（如果存在同名公共同义词则不创建）"
            Case 5
                lblNote.Caption = "    根据当前系统的转出表定义，检查历史库与在线库在表、列、约束、索引等方面的一致性，并提供修复功能。"
            Case 6
                lblNote.Caption = "    对管理工具的对象进行重新授权并创建同义词。"
            End Select
        End If
    Next
End Sub

Private Sub cmdGetIni_Click()
    Dim varData As Variant
    Dim strToolsByServer As String
    Dim strToolsVer As String
    
    With frmMDIMain.dlgMain
        .FileName = lblFileName.Caption
        .DialogTitle = "选择应用安装配置文件"
        If cmbSystem.Tag = "ZLTOOLS" Then
            .Filter = "(服务器工具脚本)|zlServer.Sql"
        Else
            .Filter = "(应用安装配置文件)|zlSetup.ini"
        End If
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            lblFileName.Caption = .FileName
        End If
    End With
    '设置功能状态
    Call SetFunsState(lblFileName.Caption, True)
    
End Sub

Private Sub Form_Load()
    
    lblMain.Caption = "通常对象检查修正的操作都需要较长时间的操作，除非系统已经明确地暴露出问题，请不要随意使用此功能。" & _
        vbCrLf & vbCrLf & "对象检查修正的所有功能都涉及对象的独占操作，如果不能确认已经没有其他用户在使用系统，千万不要运行，以免对系统造成破坏（建议断开除本机以外的所有网络连接）。"
        
    Call LoadSystem
End Sub
Private Function LoadSystem() As Boolean
    '-------------------------------------------------------------------------------------------------
    '功能:加载系统数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/10
    '-------------------------------------------------------------------------------------------------
    Dim strToolsVer As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    
    LoadSystem = False
    err = 0: On Error GoTo errHand:
    LoadSystem = True
    
    Set rsTemp = OpenCursor(gcnOracle, "zlTools.B_Public.Get_Ver")
    strToolsVer = rsTemp!内容
    '填写已安装系统清单
    If gblnDBA Then
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", gstrUserName)
    End If

    With rsTemp
        Do While Not .EOF
            cmbSystem.AddItem !名称 & " v" & !版本号 & "（" & !编号 & "）"
            cmbSystem.ItemData(cmbSystem.NewIndex) = !编号
            .MoveNext
        Loop
        cmbSystem.AddItem "服务器管理工具" & " v" & strToolsVer & "（ZLTOOLS）"
        cmbSystem.ItemData(cmbSystem.NewIndex) = -1
        If cmbSystem.ListCount = 0 Then
            cmdGetIni.Enabled = False
            For i = 0 To cmdFunction.Ubound
                cmdFunction(i).Enabled = False
            Next
        End If
        If cmbSystem.ListCount > 0 Then cmbSystem.ListIndex = 0
        If cmbSystem.ListCount = 1 Then cmbSystem.Locked = True
    End With
    '加载管理工具
    Exit Function
errHand:
    cmbSystem.AddItem "服务器管理工具" & " v" & strToolsVer & "（ZLTOOLS）"
    cmbSystem.ItemData(cmbSystem.NewIndex) = -1
    MsgBox err.Description, vbCritical, Me.Caption
End Function
Private Sub Form_Resize()
    On Error Resume Next
    Dim sngWidth As Long '最小宽度
    
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
        
    With lblMain
        .Top = lblNote.Top + lblNote.Height + 300
        .Height = ScaleHeight - picStatus.Height - .Top - 100
        .Left = lblFileName.Left
        .Width = ScaleWidth - .Left - imgMain.Left
    End With
    
    sngWidth = IIf(ScaleWidth < 5600, 5600, ScaleWidth)
    cmbSystem.Width = sngWidth - cmbSystem.Left - 300
    cmdGetIni.Left = cmbSystem.Left + cmbSystem.Width - cmdGetIni.Width
    lblFileName.Width = sngWidth - lblFileName.Left - 300
    lblNote.Width = sngWidth - lblNote.Left - 300
    
End Sub


Private Function CheckIniFile(FileName As String, Optional blnMsg As Boolean) As Boolean
    Dim strTemp As String
    Dim objText As TextStream
    Dim intDefSysCode As Integer                '系统编号
    Dim strDefSysName As String                  '系统名称
    Dim strDefVersion As String                 '版本号
    Dim strDefSpace   As String                 '表空间
    
    err = 0
    On Error Resume Next
    
    mstrIniPath = Mid(FileName, 1, Len(FileName) - 11)
    '相关文件匹配性检查
    strTemp = ""
    If Dir(mstrIniPath & "zlSequence.sql") = "" Then strTemp = strTemp & vbCr & "序列文件" & mstrIniPath & "zlSequence.sql"
    If Dir(mstrIniPath & "zlTable.sql") = "" Then strTemp = strTemp & vbCr & "数据表文件" & mstrIniPath & "zlTable.sql"
    If Dir(mstrIniPath & "zlConstraint.sql") = "" Then strTemp = strTemp & vbCr & "约束文件" & mstrIniPath & "zlConstraint.sql"
    If Dir(mstrIniPath & "zlIndex.sql") = "" Then strTemp = strTemp & vbCr & "索引文件" & mstrIniPath & "zlIndex.sql"
    If Dir(mstrIniPath & "zlView.sql") = "" Then strTemp = strTemp & vbCr & "视图文件" & mstrIniPath & "zlView.sql"
    If Dir(mstrIniPath & "zlProgram.sql") = "" Then strTemp = strTemp & vbCr & "函数过程文件" & mstrIniPath & "zlProgram.sql"
    
    '不检查,因为9系统没有此文件
    'If Dir(mstrIniPath & "zlPackage.sql") = "" Then strTemp = strTemp & vbCr & "包文件" & mstrIniPath & "zlPackage.sql"
    
    If Dir(mstrIniPath & "zlManData.sql") = "" Then strTemp = strTemp & vbCr & "管理数据文件" & mstrIniPath & "zlManData.sql"
    If Dir(mstrIniPath & "zlAppData.sql") = "" Then strTemp = strTemp & vbCr & "应用数据文件" & mstrIniPath & "zlAppData.sql"
    If strTemp <> "" Then
        If blnMsg Then MsgBox "以下服务器安装的相关文件丢失，不能继续，包括：" & strTemp, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '配置文件正确性检查
    Set objText = gobjFile.OpenTextFile(FileName)
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
    objText.Close
    
    If err <> 0 Then
        CheckIniFile = False
        If blnMsg Then MsgBox "安装配置文件不正确", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '配置文件符合性检查
    If intDefSysCode <> cmbSystem.ItemData(cmbSystem.ListIndex) \ 100 Then
        err.Raise 10
        If blnMsg Then MsgBox "选择文件不是该系统的安装配置文件", vbExclamation, gstrSysName
    ElseIf InStr(1, cmbSystem.Text, Trim(strDefVersion)) = 0 Then
        err.Raise 10
        If blnMsg Then MsgBox "选择文件与该系统版本不符", vbExclamation, gstrSysName
    End If
    If err = 0 Then
        CheckIniFile = True
    Else
        CheckIniFile = False
    End If
End Function

Private Sub InputErrRpt(ObjType As String, ObjName As String, ErrInfo As String, Optional Advice As String)
    '----------------------------------------------------
    '填写一行错误报告
    '----------------------------------------------------
    With frmAppChkRpt.hgdReport
        .Rows = .Rows + 1
        If .Tag <> ObjType Then
            .TextMatrix(.Rows - 1, 0) = "------< " & ObjType & "检查情况 >------"
            .TextMatrix(.Rows - 1, 1) = "------< " & ObjType & "检查情况 >------"
            .MergeRow(.Rows - 1) = True
            .Rows = .Rows + 1
        End If
        .Tag = ObjType
        .TextMatrix(.Rows - 1, 0) = ObjName
        If .ColData(0) < Me.TextWidth(ObjName) Then
            .ColData(0) = Me.TextWidth(ObjName)
        End If
        .TextMatrix(.Rows - 1, 1) = ErrInfo
        If .ColData(1) < Me.TextWidth(ErrInfo) Then
            .ColData(1) = Me.TextWidth(ErrInfo)
        End If
        .TextMatrix(.Rows - 1, 2) = Advice
        If .ColData(2) < Me.TextWidth(Advice) Then
            .ColData(2) = Me.TextWidth(Advice)
        End If
            
    End With
        
    '显示提醒标签
    If InStr(Advice, "严重") > 0 Or InStr(Advice, "较重") > 0 Then
         frmAppChkRpt.lblWarn.Visible = True
    End If
End Sub


Private Sub InputErrModifyRpt(ObjType As String, ObjName As String, strTableName As String, strErrType As String, ErrInfo As String, strModifyInfor As String, Optional strModifyErrInfor As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:书写指定行的对象修正情况
    '入参:
    '出参:
    '返回:
    '问题:22507
    '编制:刘兴洪
    '日期:2009-08-26 14:34:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With frmAppChkRpt.hgdReport
        .Rows = .Rows + 1
        If .Tag <> ObjType Then
            .TextMatrix(.Rows - 1, 0) = "------< " & ObjType & " >------"
            .TextMatrix(.Rows - 1, 1) = "------< " & ObjType & " >------"
            .TextMatrix(.Rows - 1, 2) = "------< " & ObjType & " >------"
            .TextMatrix(.Rows - 1, 3) = "------< " & ObjType & " >------"
            .TextMatrix(.Rows - 1, 4) = "------< " & ObjType & " >------"
            .TextMatrix(.Rows - 1, 5) = "------< " & ObjType & " >------"
            .MergeRow(.Rows - 1) = True
            .Rows = .Rows + 1
        End If
        .Tag = ObjType: i = 0
        .TextMatrix(.Rows - 1, i) = ObjName
        If .ColData(i) < Me.TextWidth(ObjName) Then .ColData(i) = Me.TextWidth(ObjName)
        
        i = i + 1: .TextMatrix(.Rows - 1, i) = strTableName
        If .ColData(i) < Me.TextWidth(strTableName) Then .ColData(i) = Me.TextWidth(strTableName)
        
        
        i = i + 1: .TextMatrix(.Rows - 1, i) = strErrType
        If .ColData(i) < Me.TextWidth(strErrType) Then .ColData(i) = Me.TextWidth(strErrType)
        
        
        i = i + 1: .TextMatrix(.Rows - 1, i) = ErrInfo
        If .ColData(i) < Me.TextWidth(ErrInfo) Then .ColData(i) = Me.TextWidth(ErrInfo)
        
        i = i + 1: .TextMatrix(.Rows - 1, i) = strModifyInfor
        If .ColData(i) < Me.TextWidth(strModifyInfor) Then .ColData(i) = Me.TextWidth(strModifyInfor)
            
        i = i + 1: .TextMatrix(.Rows - 1, i) = strModifyErrInfor
        If .ColData(i) < Me.TextWidth(strModifyErrInfor) Then .ColData(i) = Me.TextWidth(strModifyErrInfor)
        
    End With
End Sub

Private Sub CheckTable(FileName As String, Optional ByVal blnAddToRsTable As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据表，同时判断数据表的列是否正确
    '入参:FileName-脚本文件名(包含具体路径)
    '     blnAddToRsTable-是否将相关的错误信息写在记录集中
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-08-19 15:03:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str原字段精度 As String, str现字段精度 As String, strSQL1 As String
    Dim arySql() As String, strObjName As String
    Dim intVer As Integer
    Dim rsObjects As New ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim rsColumns As New ADODB.Recordset
    
    intVer = GetOracleVersion

    With rsObjects
        .Filter = 0
        If gblnDBA Then
            strSQL = "select TABLE_NAME from DBA_TABLES where OWNER='" & cmbSystem.Tag & "' And Instr(Table_Name, 'BIN$') <= 0"
        Else
            strSQL = "select TABLE_NAME from USER_TABLES where Instr(Table_Name, 'BIN$') <= 0"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        
        err = 0
        On Error Resume Next
        If Not mclsRunScript.OpenFile(FileName) Then Exit Sub
        On Error GoTo 0
        Do While Not mclsRunScript.EOF
            strSQL = UCase(TrimEx(mclsRunScript.SQLInfo.SQL))
            strSQL1 = strSQL
            arySql = Split(strSQL, " TABLE ")
            strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
            If InStr(1, strSQL, " ") > 0 And InStr(strSQL, " ") < InStr(strSQL, "(") Then
                strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
            Else
                strObjName = Trim(Left(strSQL, InStr(strSQL, "(") - 1))
            End If
            strObjName = Replace(strObjName, vbCrLf, "")
            .Filter = "TABLE_NAME='" & strObjName & "'"
            
            If .EOF Then
                If blnAddToRsTable Then
                    '1-存在对象,2-不存在对象,3-失效
                     Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "数据表", 2, False, strSQL1, "不存在", "严重：部分功能不能正常运行")
                Else
                    Call InputErrRpt("数据表", strObjName, "不存在", "严重：部分功能不能正常运行")
                End If
            Else
                '通过建立一个标准结构替代表，进行字段结构分析判断
                strSQL = arySql(0) & " table CK" & Trim(arySql(1))
                strSQL = Split(strSQL, "TABLESPACE")(0)
                On Error Resume Next
                'gcnOracle.Execute "drop table CK" & strObjName & IIf(intVer >= 100, " Purge", "")
                gcnOldOra.Execute "drop table CK" & strObjName & IIf(intVer >= 100, " Purge", "")
                err = 0
                err.Clear
                '索引组织表,必须在创建表的同时创建主键，所以需做特殊处理
                If InStr(UCase(strSQL), "PRIMARY") > 0 Then
                    strSQL = Replace(strSQL, strObjName & "_PK", "CK" & strObjName & "_PK")
                End If
                
                gcnOldOra.Execute strSQL
                If err = 0 Then
                    With rsColumns
                        If gblnDBA Then
                            strTemp = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
                                    " From DBA_TAB_COLUMNS" & _
                                    " WHERE OWNER='" & cmbSystem.Tag & "' and TABLE_NAME='" & strObjName & "'"
                        Else
                            strTemp = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
                                "       From USER_TAB_COLUMNS" & _
                                "       WHERE TABLE_NAME='" & strObjName & "'"
                        End If
                        strTemp = "select N.COLUMN_NAME as N_NAME,N.DATA_TYPE as N_TYPE,N.DATA_LENGTH as N_NLENGTH," & _
                                "        N.DATA_PRECISION as N_PRECISION,N.DATA_SCALE as N_SCALE,N.DATA_DEFAULT as N_DEFAULT," & _
                                "        O.COLUMN_NAME as O_NAME,O.DATA_TYPE as O_TYPE,O.DATA_LENGTH as O_NLENGTH," & _
                                "        O.DATA_PRECISION as O_PRECISION,O.DATA_SCALE as O_SCALE,O.DATA_DEFAULT as O_DEFAULT" & _
                                " from (SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
                                "       From USER_TAB_COLUMNS" & _
                                "       WHERE TABLE_NAME='CK" & strObjName & "') N," & _
                                "      (" & strTemp & ") O" & _
                                " where N.COLUMN_NAME=O.COLUMN_NAME(+) "    'and N.DATA_TYPE=O.DATA_TYPE(+):刘兴宏:2007/06/30可能存在字段类型发生变化的情况,因此需要取消这句
                        If .State = adStateOpen Then .Close
                        .Open strTemp, gcnOracle, adOpenKeyset
                        
                        Do While Not .EOF
                            str原字段精度 = "": str现字段精度 = ""
                            If IsNull(!O_TYPE) Then
                                '缺少的字段
                                Select Case !N_TYPE
                                Case "NUMBER"
                                    strTemp = !N_NAME & " NUMBER(" & !N_PRECISION & "," & !N_SCALE & ")"
                                    If Not IsNull(!N_DEFAULT) Then strTemp = strTemp & " DEFAULT " & !N_DEFAULT
                                    str现字段精度 = !N_PRECISION & "," & !N_SCALE
                                Case "VARCHAR2"
                                    strTemp = !N_NAME & " VARCHAR2(" & !N_NLENGTH & ")"
                                    If Not IsNull(!N_DEFAULT) Then strTemp = strTemp & " DEFAULT " & !N_DEFAULT
                                     str现字段精度 = Nvl(!N_NLENGTH)
                                Case Else
                                    strTemp = !N_NAME & !N_TYPE
                                End Select
                                If blnAddToRsTable Then
                                    '1-存在对象,2-不存在对象,3-失效,4-缺少列
                                     strSQL1 = " Alter Table " & strObjName & " Add(" & strTemp & ")"
                                     Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "数据表", 4, False, strSQL1, "缺少列 " & strTemp, "严重：部分功能不能正常运行", Nvl(!N_NAME), "", Nvl(!N_TYPE), str原字段精度, str现字段精度)
                                Else
                                    Call InputErrRpt("数据表", strObjName, "缺少列 " & strTemp, "严重：部分功能不能正常运行")
                                End If
                            Else
                                '长度不等
                                Select Case !N_TYPE
                                Case "NUMBER"
                                    If !N_PRECISION > !O_PRECISION Or !N_SCALE > !O_SCALE Then
                                        strTemp = !N_NAME & "列长度小于规定值：应为“" & "NUMBER(" & !N_PRECISION & "," & !N_SCALE & ")”" & _
                                                 " 现为“" & "NUMBER(" & !O_PRECISION & "," & !O_SCALE & ")”"
                                                 
                                        str原字段精度 = !O_PRECISION & "," & !O_SCALE: str现字段精度 = !N_PRECISION & "," & !N_SCALE
                                        If blnAddToRsTable Then
                                            If Nvl(!O_TYPE) = Nvl(!N_TYPE) Then
                                                '类型相同的情况下，才处理
                                                '1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度问题
                                                 strSQL1 = " Alter Table " & strObjName & " Modify(" & !N_NAME & " NUMBER(" & !N_PRECISION & IIf(Val(Nvl(!N_SCALE)) = 0, "", "," & Nvl(!N_SCALE) & ")") & "))"
                                                If !N_PRECISION > !O_PRECISION Then
                                                    Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "数据表", 5, False, strSQL1, "精度过小 ", "严重：较大的数据将无法正确存储：" & strTemp, Nvl(!N_NAME), Nvl(!N_TYPE), Nvl(!N_TYPE), str原字段精度, str现字段精度)
                                                ElseIf !N_SCALE > !O_SCALE Then
                                                    Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "数据表", 5, False, strSQL1, "精度过小 ", "严重：可能导致数据精度不足：" & strTemp, Nvl(!N_NAME), Nvl(!N_TYPE), Nvl(!N_TYPE), str原字段精度, str现字段精度)
                                                Else
                                                    Call InputErrRpt("数据表", strObjName, strTemp, "较轻：基本不影响运行")
                                                    Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "数据表", 5, False, strSQL1, "精度过大 ", "较轻：基本不影响运行：" & strTemp, Nvl(!N_NAME), Nvl(!N_TYPE), Nvl(!N_TYPE), str原字段精度, str现字段精度)
                                                End If
                                            End If
                                        Else
                                            If !N_PRECISION > !O_PRECISION Then
                                                Call InputErrRpt("数据表", strObjName, strTemp, "较重：较大的数据将无法正确存储")
                                            ElseIf !N_SCALE > !O_SCALE Then
                                                Call InputErrRpt("数据表", strObjName, strTemp, "较重：可能导致数据精度不足")
                                            Else
                                                Call InputErrRpt("数据表", strObjName, strTemp, "较轻：基本不影响运行")
                                            End If
                                        End If
                                                 
                                    End If
                                Case "VARCHAR2"
                                    If !N_NLENGTH <> !O_NLENGTH Then
                                        strTemp = !N_NAME & "列长度小于规定值：应为“" & "VARCHAR2(" & !N_NLENGTH & ")”" & _
                                                 " 现为“" & "VARCHAR2(" & !O_NLENGTH & ")”"
                                        str原字段精度 = !O_NLENGTH: str现字段精度 = !N_NLENGTH
                                        If blnAddToRsTable Then
                                            If Nvl(!O_TYPE) = Nvl(!N_TYPE) Then
                                                '类型相同的情况下，才处理
                                                '1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度问题
                                                 strSQL1 = " Alter Table " & strObjName & " Modify(" & !N_NAME & " VARCHAR2(" & !N_NLENGTH & ")" & ")"
                                                
                                                If !N_NLENGTH > !O_NLENGTH Then
                                                    Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "数据表", 5, False, strSQL1, "精度过小", "较重：可能导致较长文本无法存储：" & strTemp, Nvl(!N_NAME), Nvl(!N_TYPE), Nvl(!N_TYPE), str原字段精度, str现字段精度)
                                                Else
                                                    Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "数据表", 5, False, strSQL1, "精度过大", "较轻：基本不影响运行：" & strTemp, Nvl(!N_NAME), Nvl(!N_TYPE), Nvl(!N_TYPE), str原字段精度, str现字段精度)
                                                End If
                                            End If
                                        Else
                                            If !N_NLENGTH > !O_NLENGTH Then
                                                Call InputErrRpt("数据表", strObjName, strTemp, "较重：可能导致较长文本无法存储")
                                            Else
                                                Call InputErrRpt("数据表", strObjName, strTemp, "较轻：基本不影响运行")
                                            End If
                                        End If
                                    End If
                                Case Else
                                End Select
                                '刘兴宏:2007/06/30加入的列类型判断
                                If Nvl(!O_TYPE) <> Nvl(!N_TYPE) Then
                                     strTemp = !N_NAME & "列类型不一至：应为“" & Nvl(!N_TYPE) & "” 现为：“" & Nvl(!O_TYPE) & "”"
                                     If blnAddToRsTable Then
                                        Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "数据表", 5, False, "", "字段类型不对", "较重：可能导致数据不能存储!" & strTemp, Nvl(!N_NAME), Nvl(!O_TYPE), Nvl(!N_TYPE), "", "")
                                     Else
                                        Call InputErrRpt("数据表", strObjName, strTemp, "较重：可能导致数据不能存储!")
                                     End If
                                End If
                            End If
                            
                            .MoveNext
                        Loop
                    End With
                    gcnOracle.Execute "drop table CK" & strObjName & IIf(intVer >= 100, " Purge", "")
                Else
                    '创建CK表失败
                    '加入日志,提醒用户,该表未检查
                    If blnAddToRsTable = False Then
                        Call InputErrRpt("数据表", strObjName, "创建比较对象失败", "较重：不能对此对象的正确性做出判断！")
                    End If
                End If
            End If
            
            pgbState.value = mclsRunScript.Line / mclsRunScript.LinesCount * 100
            Call mclsRunScript.ReadNextSQL
            DoEvents
        Loop
    End With
    pgbState.value = 0
End Sub

Private Sub CheckConstraint(FileName As String, Optional ByVal blnAddToRsTable As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查对象约束，判断是否有效存在
    '入参:FileName-脚本文件名(包含具体路径)
    '     blnAddToRsTable-是否将相关的错误信息写在记录集中
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-08-19 15:03:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsObjects As New ADODB.Recordset, strSQL As String
    Dim rsColumns As New ADODB.Recordset, strTemp As String
    
    
    Dim str原字段精度 As String, str现字段精度 As String, strSQL1 As String, strTemp1 As String
    Dim arySql() As String, strObjName As String, strColumns As String, strTableName As String
    With rsObjects
        .Filter = 0
        If gblnDBA Then
            strSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD,Search_Condition from DBA_CONSTRAINTS where OWNER='" & cmbSystem.Tag & "'"
        Else
            strSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD,Search_Condition from USER_CONSTRAINTS"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
    
        err = 0
        On Error Resume Next
        If Not mclsRunScript.OpenFile(FileName) Then Exit Sub
        On Error GoTo 0
        Do While Not mclsRunScript.EOF
            strSQL = mclsRunScript.SQLInfo.PartSQL
            strSQL1 = strSQL
            arySql = Split(strSQL, " CONSTRAINT ")
            If UBound(arySql) > 0 Then
                '刘兴宏加入:
                strTableName = Trim(Split(Trim(arySql(0)), "TABLE")(1))
                strTableName = Split(strTableName, " ")(0)
                strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
                strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
                .Filter = "CONSTRAINT_NAME='" & strObjName & "'"
                If .EOF Then
                    If blnAddToRsTable Then
                        '1-存在对象,2-不存在对象,3-失效
                        '类型:主键,唯一,外键,约束,索引,视图,...
                        If InStr(1, UCase(Replace(strSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "外键", 2, False, strSQL1, "不存在", "较重：可能导致数据不一致，影响运行速度")
                        ElseIf InStr(1, strSQL, " CHECK") > 0 Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 2, False, strSQL1, "不存在", "较轻：基本不影响系统运行")
                        Else
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 2, False, strSQL1, "不存在", "较重：可能导致数据不一致，影响运行速度")
                        End If
                    Else
                        If InStr(1, strSQL, " CHECK") > 0 Then
                            Call InputErrRpt("约束", strObjName, "不存在", "较轻：基本不影响系统运行")
                        Else
                            Call InputErrRpt("约束", strObjName, "不存在", "较重：可能导致数据不一致，影响运行速度")
                        End If
                    End If
                ElseIf .Fields("STATUS").value <> "ENABLED" Then
                    If blnAddToRsTable Then
                        '状态:1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态
                        '类型:主键,唯一,外键,约束,索引,视图,...
                        If InStr(1, UCase(Replace(strSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "外键", 6, False, strSQL1, "当前处于禁止状态", "较重：可能系统已经存在问题")
                        Else
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 6, False, strSQL1, "当前处于禁止状态", "较重：可能系统已经存在问题")
                        End If
                    Else
                        Call InputErrRpt("约束", strObjName, "当前处于禁止状态", "较重：可能系统已经存在问题")
                    End If
                ElseIf !VALIDATED <> "VALIDATED" Then
                    If blnAddToRsTable Then
                        '状态:1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态
                        '类型:主键,唯一,外键,约束,索引,视图,...
                        If InStr(1, UCase(Replace(strSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "外键", 3, False, strSQL1, "当前处于无效状态", "较重：可能数据一致性已被破坏")
                        Else
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 3, False, strSQL1, "当前处于无效状态", "较重：可能数据一致性已被破坏")
                        End If
                    Else
                        Call InputErrRpt("约束", strObjName, "当前处于无效状态", "较重：可能数据一致性已被破坏")
                    End If
                ElseIf Not IsNull(!BAD) Then
                    If blnAddToRsTable Then
                        '状态:1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态(损坏)
                        '类型:主键,唯一,外键,约束,索引,视图,...
                        If InStr(1, UCase(Replace(strSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "外键", 6, False, strSQL1, "约束被意外损坏", "严重：可能存在硬件错误")
                        Else
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 6, False, strSQL1, "约束被意外损坏", "严重：可能存在硬件错误")
                        End If
                    Else
                        Call InputErrRpt("约束", strObjName, "约束被意外损坏", "严重：可能存在硬件错误")
                    End If
                Else
                    strColumns = ""
                    With rsColumns
                        If gblnDBA Then
                            strTemp = "select COLUMN_NAME" & _
                                " from DBA_CONS_COLUMNS" & _
                                " where OWNER='" & cmbSystem.Tag & "' and CONSTRAINT_NAME='" & strObjName & "'" & _
                                " order by POSITION"
                        Else
                            strTemp = "select COLUMN_NAME" & _
                                " from USER_CONS_COLUMNS" & _
                                " where CONSTRAINT_NAME='" & strObjName & "'" & _
                                " order by POSITION"
                        End If
                        If .State = adStateOpen Then .Close
                        .Open strTemp, gcnOracle, adOpenKeyset
                        Do While Not .EOF
                            strColumns = strColumns & "," & !Column_Name
                            .MoveNext
                        Loop
                    End With
                    
                    If InStr(1, strSQL, " PRIMARY ") > 0 Then
                        If !constraint_type <> "P" Then
                            If blnAddToRsTable Then
                                '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                                '类型:主键,唯一,外键,约束,索引,视图,...
                                Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 7, False, strSQL1, "约束类型错误", "严重：可能影响系统运行 ,应为主键约束")
                                If !constraint_type = "U" Then
                                    '同时，需要检查相关的级联外键
                                    Call zl获取关联外键(strObjName, mrsErrTable)
                                End If
                            Else
                                Call InputErrRpt("约束", strObjName, "约束类型错误，应为主键约束", "严重：可能影响系统运行")
                            End If
                        Else
                            arySql = Split(strSQL, " PRIMARY ")
                            strTemp = Replace(Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "KEY", ""), "(", ""), " ", "")
                            If strColumns <> "," & strTemp Then
                                If blnAddToRsTable Then
                                    '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                                    '类型:主键,唯一,外键,约束,索引,视图,...
                                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 4, False, strSQL1, "约束列错误", "较重：可能影响系统运行，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")")
                                    '同时，需要检查相关的级联外键
                                    Call zl获取关联外键(strObjName, mrsErrTable)
                                    
                                Else
                                    Call InputErrRpt("约束", strObjName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "较重：可能影响系统运行")
                                End If
                            End If
                        End If
                    ElseIf InStr(1, strSQL, " UNIQUE") > 0 Then
                        If !constraint_type <> "U" Then
                            If blnAddToRsTable Then
                                '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                                '类型:主键,唯一,外键,约束,索引,视图,...
                                Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 7, False, strSQL1, "约束类型错误", "较重：可能影响系统运行 ,应为唯一约束")
                                If !constraint_type = "P" Then
                                    '同时，需要检查相关的级联外键
                                    Call zl获取关联外键(strObjName, mrsErrTable)
                                End If
                            Else
                                Call InputErrRpt("约束", strObjName, "约束类型错误，应为唯一约束", "较重：可能影响系统运行")
                            End If
                        Else
                            arySql = Split(strSQL, " UNIQUE ")
                            If UBound(arySql) = 0 Then arySql = Split(strSQL, " UNIQUE(")
                            strTemp = Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "(", ""), " ", "")
                            If strColumns <> "," & strTemp Then
                                If blnAddToRsTable Then
                                    '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                                    '类型:主键,唯一,外键,约束,索引,视图,...
                                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 4, False, strSQL1, "约束列错误", "较重：可能影响系统运行，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")")
                                    '同时，需要检查相关的级联外键
                                    Call zl获取关联外键(strObjName, mrsErrTable)
                                Else
                                    Call InputErrRpt("约束", strObjName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "较重：可能影响系统运行")
                                End If
                            End If
                        End If
                    ElseIf InStr(1, strSQL, " FOREIGN ") > 0 Then
                        If !constraint_type <> "R" Then
                            If blnAddToRsTable Then
                                '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                                '类型:主键,唯一,外键,约束,索引,视图,...
                                Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 7, False, strSQL1, "约束类型错误", "较重：可能影响系统运行，应为外键约束")
                            Else
                                Call InputErrRpt("约束", strObjName, "约束类型错误，应为外键约束", "严重：可能影响系统运行")
                            End If
                        Else
                            arySql = Split(strSQL, " FOREIGN ")
                            strTemp = Replace(Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "KEY", ""), "(", ""), " ", "")
                            If strColumns <> "," & strTemp Then
                                If blnAddToRsTable Then
                                    '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                                    '类型:主键,唯一,外键,约束,索引,视图,...
                                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 4, False, strSQL1, "约束列错误", "较重：可能影响系统运行，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")")
                                Else
                                    Call InputErrRpt("约束", strObjName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "严重：可能影响数据一致性")
                                End If
                            End If
                        End If
                    ElseIf InStr(1, strSQL, " CHECK") > 0 Then
                        If !constraint_type <> "C" Then
                            If blnAddToRsTable Then
                                '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                                '类型:主键,唯一,外键,约束,索引,视图,...
                                Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 7, False, strSQL1, "约束类型错误", "较重：可能影响系统运行，应为检查约束")
                            Else
                                Call InputErrRpt("约束", strObjName, "约束类型错误，应为检查约束", "严重：可能影响系统运行")
                            End If
                        Else
                            '25047:刘兴洪增加
                            arySql = Split(strSQL, " CHECK")
                            strTemp = Replace(UCase(Replace(Replace(Replace(arySql(1), " ", ""), vbTab, ""), vbCrLf, "")), ";", "")
                            strTemp1 = "(" & Replace(UCase(Replace(Replace(Replace(Nvl(!Search_Condition), " ", ""), vbTab, ""), vbCrLf, "")), ";", "") & ")"
                            If strTemp <> strTemp1 Then
                                '检查约束检查是否一致
                                strTemp = Trim(Replace(Replace(Replace(arySql(1), vbCrLf, " "), vbTab, " "), ";", ""))
                                strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
                                
                                If blnAddToRsTable Then
                                    '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                                    '类型:主键,唯一,外键,约束,索引,视图,...
                                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 7, False, strSQL1, "约束内容不一致", "较重：可能影响系统运行,导致数据不一致!,应为(" & strTemp & "),现在为(" & Nvl(!Search_Condition) & ")")
                                Else
                                    Call InputErrRpt("约束", strObjName, "约束内容不一致,应为(" & strTemp & "),现在为(" & Nvl(!Search_Condition) & ")", "较重：可能影响系统运行,导致数据不一致!")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            pgbState.value = mclsRunScript.Line / mclsRunScript.LinesCount * 100
            Call mclsRunScript.ReadNextSQL
            DoEvents
        Loop
    End With
    pgbState.value = 0
End Sub
Private Sub CheckIndex(FileName As String, Optional ByVal blnAddToRsTable As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查对象约束，判断是否有效存在
    '入参:FileName-脚本文件名(包含具体路径)
    '     blnAddToRsTable-是否将相关的错误信息写在记录集中
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-08-19 15:03:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str原字段精度 As String, str现字段精度 As String, strSQL1 As String
    Dim arySql() As String, strObjName As String, strColumns As String
    Dim strTablenName As String
    Dim rsObjects As New ADODB.Recordset, strSQL As String
    Dim rsColumns As New ADODB.Recordset, strTemp As String
    With rsObjects
        .Filter = 0
        If gblnDBA Then
            strSQL = "select INDEX_NAME,STATUS from DBA_INDEXES where OWNER='" & cmbSystem.Tag & "' and TABLE_OWNER='" & cmbSystem.Tag & "'"
        Else
            strSQL = "select INDEX_NAME,STATUS from USER_INDEXES where TABLE_OWNER='" & cmbSystem.Tag & "'"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
    
        err = 0
        On Error Resume Next
        If Not mclsRunScript.OpenFile(FileName) Then Exit Sub
        On Error GoTo 0
        Do While Not mclsRunScript.EOF
            strSQL = mclsRunScript.SQLInfo.PartSQL
            strSQL1 = strSQL
            arySql = Split(strSQL, " INDEX ")
            If UBound(arySql) > 0 Then
                strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
                strTablenName = Trim(Split(Split(strSQL, "ON")(1), "(")(0))
                strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
                .Filter = "INDEX_NAME='" & strObjName & "'"
                If .EOF Then
                    If blnAddToRsTable Then
                        '状态:1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态
                        '类型:主键,唯一,外键,约束,索引,视图,...
                        Call zlInsertRecData(mrsErrTable, strTablenName, strObjName, "索引", 2, False, strSQL1, "不存在", "较重：可能影响系统运行速度")
                    Else
                        Call InputErrRpt("索引", strObjName, "不存在", "较重：可能影响系统运行速度")
                    End If
                ElseIf .Fields("STATUS").value <> "VALID" Then
                    Call InputErrRpt("索引", strObjName, "当前处于无效状态")
                Else
                    With rsColumns
                        If gblnDBA Then
                            strTemp = "select TABLE_NAME,COLUMN_NAME" & _
                                    " from DBA_IND_COLUMNS" & _
                                    " where INDEX_OWNER='" & cmbSystem.Tag & "' and INDEX_NAME='" & strObjName & "'" & _
                                    " order by COLUMN_POSITION"
                        Else
                            strTemp = "select TABLE_NAME,COLUMN_NAME" & _
                                    " from USER_IND_COLUMNS" & _
                                    " where INDEX_NAME='" & strObjName & "'" & _
                                    " order by COLUMN_POSITION"
                        End If
                        If .State = adStateOpen Then .Close
                        .Open strTemp, gcnOracle, adOpenKeyset
                        Do While Not .EOF
                            If .AbsolutePosition = 1 Then
                                strColumns = !Table_Name & "(" & !Column_Name
                            Else
                                strColumns = strColumns & "," & !Column_Name
                            End If
                            .MoveNext
                        Loop
                            strColumns = strColumns & ")"
                    End With
                    arySql = Split(strSQL, " ON ")
                    strTemp = Replace(Left(arySql(1), InStr(1, arySql(1), ")")), " ", "")
                    If strColumns <> strTemp Then
                        If blnAddToRsTable Then
                            '状态:1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态
                            '类型:主键,唯一,外键,约束,索引,视图,...
                            Call zlInsertRecData(mrsErrTable, strTablenName, strObjName, "索引", 4, False, strSQL1, "索引列错误", "较重：可能影响系统运行速度,应为“" & strTemp & "”，现为“" & strColumns & "”")
                        Else
                            Call InputErrRpt("索引", strObjName, "索引列错误，应为“" & strTemp & "”，现为“" & strColumns & "”", "较重：可能影响系统运行速度")
                        End If
                    End If
                End If
            End If
            Call mclsRunScript.ReadNextSQL
        Loop
    End With
End Sub

Private Function RunSQLScript(ByVal cnThisDB As ADODB.Connection, ByVal strFile As String, Optional blnResumeNext As Boolean = True) As Boolean
'----------------------------------------------
'功能:执行SQL文件
'参数:
'    cnThisDB=当前系统连接
'    strFile=脚本文件
'    blnResumeNext=是否错误继续
'    返回：true-执行成功；false-执行失败
' ----------------------------------------------
    Dim lngLines As Long
    err = 0
    On Error Resume Next
    If Not mclsRunScript.OpenFile(strFile) Then Exit Function
    pgbState.value = 0
    Do While Not mclsRunScript.EOF
        cnThisDB.Execute mclsRunScript.SQLInfo.SQL
        If err <> 0 Then
            If blnResumeNext Then
                err.Clear
            Else
                MsgBox "由于文件" & strFile & "中存在下面错误导致执行中断：" & vbCr & mclsRunScript.SQLInfo.SQL, vbExclamation, gstrSysName
                Exit Function
            End If
        End If
        pgbState.value = mclsRunScript.Line / mclsRunScript.LinesCount * 100
        Call mclsRunScript.ReadNextSQL
        DoEvents
    Loop
    pgbState.value = 0
    RunSQLScript = True
End Function



Private Sub Form_Unload(Cancel As Integer)
    If picStatus.Visible Then Cancel = 1
    
    Set mclsObjectCheck = Nothing
    Set mclsRunScript = Nothing
    Set mrsErrTable = Nothing
    
End Sub

Private Sub mclsObjectCheck_AfterObjectCheck()
    picStatus.Visible = False
    cmdFunction(0).Enabled = True
End Sub

Private Sub mclsObjectCheck_AfterProgress()
    lblStatus.Caption = ""
    pgbState.value = 0
End Sub

Private Sub mclsObjectCheck_BeforeObjectCheck()
    picStatus.Visible = True
    cmdFunction(0).Enabled = False
End Sub

Private Sub mclsObjectCheck_BeforeProgress(ByVal Title As String, ByVal Max As Long)
    lblStatus.Caption = Title
    pgbState.Max = Max
End Sub

Private Sub mclsObjectCheck_Exception()
    Dim lngCount As Long
    
    For lngCount = 0 To cmdFunction.Ubound
        cmdFunction(lngCount).Enabled = False
    Next
    lblFileName.Caption = ""
    
    On Error Resume Next
    cmdGetIni.SetFocus
End Sub

Private Sub mclsObjectCheck_Progressing(ByVal Progress As Long)
    pgbState.value = Progress
    DoEvents
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

Private Sub zlGetConstraintInfor(FileName As String, ByRef cllOutPara As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据相关的脚本,获取相关的约束信息
    '入参:FileName-脚本文件
    '出参:cllOutPara-返回约束信息(表名,约束名,类型(UQ,PK,CK,FK),存在否(Y/N),历史数据空间修正(Y/N)))
    '返回:
    '编制:刘兴洪
    '日期:2009-07-21 15:55:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arySql() As String, strObjName As String, strColumns As String, strTableName As String
    Dim blnFK As Boolean
    Dim rsObjects As New ADODB.Recordset, strSQL As String
    
    With rsObjects
        .Filter = 0
        If gblnDBA Then
            strSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD from DBA_CONSTRAINTS where OWNER='" & cmbSystem.Tag & "'"
        Else
            strSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD from USER_CONSTRAINTS"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
    
        err = 0
        On Error Resume Next
        If Not mclsRunScript.OpenFile(FileName) Then Exit Sub
        On Error GoTo 0
        Do While Not mclsRunScript.EOF
            strSQL = mclsRunScript.FormatSQL
            arySql = Split(strSQL, " CONSTRAINT ")
            
            If UBound(arySql) > 0 Then
                '刘兴宏加入:
                strTableName = Trim(Split(Trim(arySql(0)), "TABLE")(1))
                strTableName = Split(strTableName, " ")(0)
                strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
                strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
                blnFK = InStr(1, UCase(Replace(strSQL, " ", "")), UCase("ForeignKey")) > 0
                
                
                
                
                .Filter = "CONSTRAINT_NAME='" & strObjName & "'" '
                cllOutPara.Add Array(strTableName, strObjName, IIf(blnFK, "Y", "N"), IIf(.EOF, "N", "Y"))
                
            End If
            pgbState.value = mclsRunScript.Line / mclsRunScript.LinesCount * 100
            Call mclsRunScript.ReadNextSQL
            DoEvents
        Loop
    End With
    pgbState.value = 0
End Sub
Private Sub zlGetIndexInfor(FileName As String, ByRef cllOutPara As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取索引信息
    '入参:
    '出参:cllOutPara-返回索引信息(表名,索引名,存在否(Y/N),历史数据空间修正(Y/N)))
    '返回:
    '编制:刘兴洪
    '日期:2009-07-21 17:11:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arySql() As String, strObjName As String, strColumns As String
    Dim strTablenName As String
    Dim rsObjects As New ADODB.Recordset, strSQL As String
    
    With rsObjects
        .Filter = 0
        If gblnDBA Then
            strSQL = "select INDEX_NAME,STATUS from DBA_INDEXES where OWNER='" & cmbSystem.Tag & "' and TABLE_OWNER='" & cmbSystem.Tag & "'"
        Else
            strSQL = "select INDEX_NAME,STATUS from USER_INDEXES where TABLE_OWNER='" & cmbSystem.Tag & "'"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
    
        err = 0
        On Error Resume Next
        If Not mclsRunScript.OpenFile(FileName) Then Exit Sub
        
        Do While Not mclsRunScript.EOF
            strSQL = mclsRunScript.FormatSQL
            arySql = Split(strSQL, " INDEX ")
            If UBound(arySql) > 0 Then
                strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
                strTablenName = Trim(Split(Split(strSQL, "ON")(1), "(")(0))
                strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
                .Filter = "INDEX_NAME='" & strObjName & "'"
                cllOutPara.Add Array(strTablenName, strObjName, IIf(.EOF, "N", "Y"))
            End If
            Call mclsRunScript.ReadNextSQL
        Loop
        
    End With
End Sub

Private Function CheckHistorySpaceEx(ByVal lngSys As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前系统下是否有历史表空间的表信息
    '入参:
    '出参:CheckHistorySpace-存在对应的历史表
    '返回:
    '编制:刘硕
    '日期:2013-04-07 10:25:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    
    On Error Resume Next '可能当前用户没有表权限
    gstrSQL = "Select 名称,所有者 From Zltools.Zlbakspaces Where 系统 = " & lngSys & "  And 当前 = 1 And 只读 = 0"
    Call OpenRecordset(rsTmp, gstrSQL, "读取历史表空间所有者")
    CheckHistorySpaceEx = Not rsTmp.EOF
    On Error GoTo 0
End Function

Private Function GetToolsIniVersion(ByVal strFilePath As String) As String
'功能:根据服务器管理工具脚本的安装配置文件路径获取管理工具安装脚本对应的版本。
    Dim strMaxVer As String, strVer As String
    Dim objFile As Scripting.File
    Dim intType As Integer
    
    On Error Resume Next
    For Each objFile In gobjFile.GetFile(strFilePath).ParentFolder.Files
        If AnalysisFileName(objFile.name, 0, strVer) Then
            '升级脚本的版本为大版本才进行比较
            If strVer Like "*.*.0" Then strMaxVer = IIf(VerFull(strVer) > VerFull(strMaxVer, False), strVer, strMaxVer)
        End If
    Next
    GetToolsIniVersion = strMaxVer
End Function


Private Sub SetFunsState(ByVal strFilePath As String, Optional blnMsg As Boolean)
'功能：对象检查对象修正功能可用性设置
    Dim blnTools As Boolean
    Dim blnTmp As Boolean
    Dim strToolsByServer As String
    Dim varData As Variant
    Dim strVer As String
    Dim i As Long
    
    blnTools = cmbSystem.Tag = "ZLTOOLS"
    '文件存在性检查
    blnTmp = gobjFile.FileExists(strFilePath)
    
    If Not blnTools And blnTmp Then
        blnTmp = CheckIniFile(strFilePath, blnMsg)
    End If
    
    '文件检查成功
    lblFileName.Caption = IIf(blnTmp, strFilePath, "")
    For i = 0 To cmdFunction.UBound - 1
        If i <= 1 Then
            cmdFunction(i).Enabled = (UCase(gstrUserName) = UCase(cmbSystem.Tag) And Not blnTools Or blnTools) And blnTmp
        ElseIf blnTools Then
            cmdFunction(i).Enabled = False
        ElseIf i = CMDFUN.E历史结构 Then
            cmdFunction(i).Enabled = UCase(gstrUserName) = UCase(cmbSystem.Tag) And CheckHistorySpaceEx(Val(cmbSystem.ItemData(cmbSystem.ListIndex)))
        ElseIf i = CMDFUN.E同义词 Then
            cmdFunction(i).Enabled = UCase(gstrUserName) = UCase(cmbSystem.Tag)
        Else
            cmdFunction(i).Enabled = True
        End If
    Next
    '设置对象权限修正是否可用
    cmdFunction(E对象权限).Visible = blnTools
    cmdFunction(E历史结构).Visible = Not blnTools

    If Me.Visible And Not blnTmp Then cmdGetIni.SetFocus: Exit Sub
    '版本检查
    '获取管理工具当前最大脚本的版本
    If blnTools Then strToolsByServer = GetToolsIniVersion(lblFileName.Caption)
    varData = Split(cmbSystem.Text & "v", "v")
    varData = Split(varData(1), "（")
    varData = Split(varData(0) & "....", ".")
    strVer = Val(varData(0)) & "." & Val(varData(1)) & "." & Val(varData(2))
    If Val(varData(2)) > 0 Then
        cmdFunction(0).Enabled = False
        cmdFunction(1).Enabled = False
    Else
        If blnTools And strVer <> strToolsByServer Then
            cmdFunction(0).Enabled = False
            cmdFunction(1).Enabled = False
        End If
    End If
End Sub

