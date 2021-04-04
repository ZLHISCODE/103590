VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppStart 
   BackColor       =   &H80000005&
   Caption         =   "系统装卸管理"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmAppStart.frx":0000
   ScaleHeight     =   5385
   ScaleWidth      =   9090
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdReCalc 
      Caption         =   "更新建议值(&R)"
      Height          =   350
      Left            =   3885
      TabIndex        =   11
      Top             =   3660
      Width           =   1500
   End
   Begin VB.TextBox txtMem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2100
      TabIndex        =   9
      Top             =   3465
      Width           =   630
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "创建(&C)…"
      Height          =   350
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   2325
      Width           =   1275
   End
   Begin MSComctlLib.ImageList imgSys 
      Left            =   4170
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppStart.frx":04F9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwSys 
      Height          =   1380
      Left            =   960
      TabIndex        =   2
      Top             =   870
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   2434
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Img大图标"
      SmallIcons      =   "imgSys"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "系统名称"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "版本号"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "编号"
         Text            =   "编号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "所有者"
         Text            =   "所有者"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ListView lvwPara 
      Height          =   1230
      Left            =   960
      TabIndex        =   4
      Top             =   4020
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   2170
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Img大图标"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "参数"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "当前值"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "建议"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "说明"
         Object.Width           =   16581
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "数据类型"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "转换为M"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "拆卸(&M)…"
      Height          =   350
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   2670
      Width           =   1275
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "再植(&R)…"
      Height          =   350
      Index           =   2
      Left            =   960
      TabIndex        =   8
      Top             =   3015
      Width           =   1275
   End
   Begin VB.Label lblMem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "服务器内存(M)       (当本机不是服务器时,请修改为服务器的物理内存大小，以给出准确的建议值。)"
      Height          =   180
      Left            =   930
      TabIndex        =   10
      Top             =   3510
      Width           =   8190
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "在取得合法的应用系统创建文件获得有效授权之后，可以创建新的系统。"
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   2310
      TabIndex        =   6
      Top             =   2340
      Width           =   2700
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPara 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "附：要求或建议的数据库参数"
      Height          =   180
      Left            =   945
      TabIndex        =   3
      Top             =   3795
      Width           =   2340
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已安装应用系统"
      Height          =   180
      Left            =   960
      TabIndex        =   1
      Top             =   675
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "系统装卸管理"
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
      Picture         =   "frmAppStart.frx":114B
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmAppStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim strSql As String
Dim objItem As ListItem
Dim intCount As Integer

Private mintVersion As Integer

Private Sub cmdFunction_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    For intCount = 0 To cmdFunction.UBound
        If intCount = Index Then
            cmdFunction(intCount).FontBold = True
            cmdFunction(intCount).SetFocus
            Select Case intCount
            Case 0
                lblNote.Caption = "    在取得合法的应用系统创建文件获得有效授权之后，可以创建新的系统。"
            Case 1
                lblNote.Caption = "    对不需用的系统，可根据安装文件进行拆卸，以降低系统的负荷。"
            Case 2
                lblNote.Caption = "    如果确认已经以其他方式(如手工执行Import)安装了应用系统的应用结构和数据，可以通过本功能植入系统管理数据。"
            End Select
        Else
            cmdFunction(intCount).FontBold = False
        End If
    Next
End Sub

Private Sub cmdFunction_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call cmdFunction_MouseMove(Index, 0, 0, 0, 0)
    Select Case Index
    Case 0
        frmAppCreate.Show 1, frmMDIMain
        Call SysCreated
    Case 1
        Dim strLinkSys As String
        If lvwSys.SelectedItem Is Nothing Then Exit Sub
        If lvwSys.SelectedItem.Selected = False Then Exit Sub
        strLinkSys = ""
        
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Loadandunload.Get_Share_name", Mid(lvwSys.SelectedItem.Key, 2))
        
        With rsTemp
            Do While Not .EOF
                strLinkSys = strLinkSys & vbCrLf & .Fields(0).value
                .MoveNext
            Loop
        End With
        If strLinkSys <> "" Then
            MsgBox "由于当前系统被以下系统共享，不能直接拆卸：" & strLinkSys, vbExclamation, gstrSysName
            Exit Sub
        End If
        frmAppRemove.Show 1, frmMDIMain
        Call SysCreated
    Case 2
        frmAppReplant.Show 1, frmMDIMain
        Call SysCreated
    End Select

End Sub

Private Sub cmdReCalc_Click()
    If IsNumeric(txtMem) Then
        If Val(txtMem) < 256 Or Val(txtMem) > 10000 Then
            MsgBox "服务器内存应在256至10000之间!", vbInformation, gstrSysName
        Else
            Call SysPara
        End If
    Else
        MsgBox "服务器内存应为整型数字!", vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_Activate()
    If Not gblnDBA Then
        frmMDIMain.stbThis.Panels(2).Text = "如需安装、拆卸应用系统，请使用获得特殊权限的DBA用户重新注册进入"
    End If
End Sub

Private Sub Form_Deactivate()
    frmMDIMain.stbThis.Panels(2).Text = ""
End Sub

Private Sub Form_Load()

    If Not gblnDBA Then
        For intCount = 0 To cmdFunction.UBound
            cmdFunction(intCount).Enabled = False
        Next
    End If
    
    mintVersion = GetOracleVersion
    '物理内存大小
    Dim mem As MEMORYSTATUS
    GlobalMemoryStatus mem
    'MsgBox "physical   Memory   is:" & mem.dwTotalPhys
    txtMem = Format(Val(mem.dwTotalPhys) / 1024 / 1024, "0")
    
    '填写系统参数
    Call SysPara

    '填写已安装系统清单
    Call SysCreated

End Sub

Private Sub Form_Resize()
    Dim sngHeight As Single
    
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    lblSys.Left = imgMain.Left + imgMain.Width + 200
    
    lvwSys.Left = lblSys.Left
    lvwSys.Width = ScaleWidth - lvwSys.Left - 200
    
    For intCount = 0 To cmdFunction.UBound
        cmdFunction(intCount).Left = lblSys.Left
    Next
    lblNote.Left = cmdFunction(0).Left + cmdFunction(0).Width + 100
    lblNote.Width = ScaleWidth - lblNote.Left - 200
    
    lblPara.Left = imgMain.Left + imgMain.Width + 200
    lvwPara.Left = lblPara.Left
    lvwPara.Width = ScaleWidth - lvwPara.Left - 200
    
    
    '设置高度
    sngHeight = IIf(ScaleHeight < 5400, 5400, ScaleHeight) '最小高度
    lvwPara.Height = 4050
    lvwPara.Top = sngHeight - lvwPara.Height - 200
    lblPara.Top = lvwPara.Top - lblPara.Height - 30
    
    cmdReCalc.Left = lvwPara.Left + lvwPara.Width - cmdReCalc.Width - 60
    cmdReCalc.Top = lvwPara.Top - cmdReCalc.Height - 15
    
    txtMem.Top = cmdReCalc.Top - lblMem.Height - 105
    txtMem.Left = lblPara.Left + 1170
    
    lblMem.Top = cmdReCalc.Top - lblMem.Height - 85
    lblMem.Left = lblPara.Left
    
    lblNote.Top = lblMem.Top - lblNote.Height - 200
    cmdFunction(0).Top = lblNote.Top - 15
    cmdFunction(1).Top = cmdFunction(0).Top + 345
    cmdFunction(2).Top = cmdFunction(1).Top + 345
    
    lvwSys.Height = lblNote.Top - lvwSys.Top - 90
    
End Sub


Private Sub SysCreated()
    lvwSys.ListItems.Clear
        
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    With rsTemp
        Do Until .EOF
            Set objItem = lvwSys.ListItems.Add(, "S" & !编号, !名称, , 1)
            objItem.SubItems(1) = IIf(IsNull(.Fields("版本号").value), "", .Fields("版本号").value)
            objItem.SubItems(2) = !编号
            objItem.SubItems(3) = IIf(IsNull(.Fields("所有者").value), "", .Fields("所有者").value)
            .MoveNext
        Loop
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMDIMain.stbThis.Panels(2).Text = ""
    mintVersion = 0
    txtMem = ""
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As zlPrintLvw
    
    Set objPrint = New zlPrintLvw
    If ActiveControl Is lvwPara Then
        objPrint.Title.Text = "要求或建议的数据库参数"
        Set objPrint.Body.objData = lvwPara
    Else
        objPrint.Title.Text = "已安装应用系统"
        Set objPrint.Body.objData = lvwSys
    End If
    objPrint.BelowAppItems.Add "打印时间：" & Format(CurrentDate, "yyyy年MM月dd日")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub SysPara()
    Dim strParas As String
    Dim blnCBO As Boolean '为真不显示 ,optimizer_index_cost_adj 和 optimizer_index_caching
    Dim blnSGA As Boolean '为真不显示 Db_cache_size,Shared_pool_size,java_pool_size
    blnCBO = True
    blnSGA = False
    
    lvwPara.ListItems.Clear
    
    With lvwPara
        
        Set objItem = lvwPara.ListItems.Add(, "open_cursors", "open_cursors")
        objItem.SubItems(2) = ">=60"
        objItem.SubItems(3) = "每个事务中可打开SQL游标的最大数量，当执行一个复杂的大事务时，可能需要打开大量的SQL游标"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "不转换"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
                
        Set objItem = lvwPara.ListItems.Add(, "session_cached_cursors", "session_cached_cursors")
        objItem.SubItems(2) = ">=10"
        objItem.SubItems(3) = "每个会话缓存的客户端游标数量,此参数影响SQL的软解析,加大参数值可提高SQL性能，但会耗用更多的服务器内存"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "不转换"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "max_enabled_roles", "max_enabled_roles")
        objItem.SubItems(2) = ">=40"
        objItem.SubItems(3) = "当需要建立较多的角色时，请修改最大角色允许数"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "不转换"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "processes", "processes")
        objItem.SubItems(2) = ">=150"
        objItem.SubItems(3) = "数据库实例的并发进程最大数量，数量过少将限制可连接的并发进程数"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "不转换"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "sessions", "sessions")
        objItem.SubItems(2) = ">=150"
        objItem.SubItems(3) = "数据库实例的并发会话最大数量，数量过少将限制可连接的并发会话数"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "不转换"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "job_queue_processes", "job_queue_processes")
        objItem.SubItems(2) = ">=10"
        objItem.SubItems(3) = "控制系统可运行的自动作业数，根据可能设置的自动作业数目设置(不超过36)"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "不转换"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        
        Set objItem = lvwPara.ListItems.Add(, "compatible", "compatible")
        objItem.SubItems(2) = ">=10.0.3"
        objItem.SubItems(3) = "兼容参数，ZLHIS标准版产品要求的最低版本为10.0.3"
        objItem.SubItems(4) = "兼容版本号"
        objItem.SubItems(5) = "不转换"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "optimizer_mode", "optimizer_mode")
        objItem.SubItems(2) = "ALL_ROWS"
        objItem.SubItems(3) = "优化器模式，建议设置为all_rows"
        objItem.SubItems(4) = "文本"
        objItem.SubItems(5) = "强调"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        '不是ALL_ROWS时才显示
        Set objItem = lvwPara.ListItems.Add(, "optimizer_index_cost_adj", "optimizer_index_cost_adj")
        objItem.SubItems(2) = "20"
        objItem.SubItems(3) = "CBO优化器模式下,计算SQL执行计划的成本时,索引相对于表扫描的成本调整比例,值越小,索引扫描的估算成本就越低"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "强调"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"

        Set objItem = lvwPara.ListItems.Add(, "optimizer_index_caching", "optimizer_index_caching")
        objItem.SubItems(2) = "80"
        objItem.SubItems(3) = "CBO优化器模式下,计算SQL执行计划的成本时,索引在内存中的估算比例,仅影响嵌套循环和in-list遍历,索引扫描的估算成本就越低"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "强调"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "cursor_sharing", "cursor_sharing")
        objItem.SubItems(2) = "EXACT"
        objItem.SubItems(3) = "此参数影响SQL的解析,建议为EXACT(精确匹配)"
        objItem.SubItems(4) = "文本"
        objItem.SubItems(5) = "强调"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        
        
        '内存配置参数
        '----------------------------------------------------------------------------------------------
        Set objItem = lvwPara.ListItems.Add(, "log_buffer", "log_buffer")
        objItem.SubItems(2) = ">=" & Val(209715200 / 1024 / 1024) & "M"
        objItem.SubItems(3) = "日志缓存区大小对于大批量数据处理（例如：升级脚本中的数据修正）影响较大，建议不低于200M，修改后需重启实例"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "转换强调"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "parallel_execution_message_size", "parallel_execution_message_size")
        objItem.SubItems(2) = ">=8192"
        objItem.SubItems(3) = "并行执行消息的大小，低于8192时，采用并行DDL重建索引时可能会报错"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "不转换"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        If mintVersion >= 90 Then
            Set objItem = lvwPara.ListItems.Add(, "db_cache_size", "db_cache_size")
            'objItem.SubItems(2) = ">=26214400"
            objItem.SubItems(2) = ">=" & Format((Val(txtMem) * 1024 * 1024 * 0.3) / 1024 / 1024, "0")
            objItem.SubItems(3) = "数据缓冲池大小(M),数据缓冲池应尽可能地大,建议设置为SGA的80%"
            objItem.SubItems(4) = "数字"
            objItem.SubItems(5) = "转换强调"
        Else
            Set objItem = lvwPara.ListItems.Add(, "db_block_buffers", "db_block_buffers")
            objItem.SubItems(2) = ">=" & Format((Val(txtMem) * 1024 * 1024 * 0.25) / 8192, "0")
            objItem.SubItems(3) = "以块大小表示的数据缓冲区,数据缓冲池应尽可能地大,建议设置为SGA的80%"
            objItem.SubItems(4) = "数字"
            objItem.SubItems(5) = "不转换"
             
        End If
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "shared_pool_size", "shared_pool_size")
        objItem.SubItems(2) = ">=" & Format((Val(txtMem) * 1024 * 1024 * 0.1) / 1024 / 1024, "0")
        objItem.SubItems(3) = "共享池(包括SQL语句缓存、系统数据字典缓存等)的内存量(M)，建议为计算内存的10-30%,共享池太大,反而影响性能"
        objItem.SubItems(4) = "数字"
        objItem.SubItems(5) = "转换强调"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
                
        '-- 9I 以上
        If mintVersion >= 90 Then
            Set objItem = lvwPara.ListItems.Add(, "workarea_size_policy", "workarea_size_policy")
            objItem.SubItems(2) = "AUTO"
            objItem.SubItems(3) = "指PGA的管理模式，建议为自动(Auto),否则需设置这些参数的合理值sort_area_size,hash_area_size,bitmap_merge_area_size"
            objItem.SubItems(4) = "文本"
            objItem.SubItems(5) = "强调"
            strParas = strParas & ",'" & Trim(objItem.Key) & "'"
            
            '
            Set objItem = lvwPara.ListItems.Add(, "pga_aggregate_target", "pga_aggregate_target")
            objItem.SubItems(2) = ">0"
            objItem.SubItems(3) = "所有会话可用的私有内存总量（每个最多可用100M以内的5%），建议分配为总的物理内存*80%*20%"
            objItem.SubItems(4) = "数字"
            objItem.SubItems(5) = "转换强调"
            strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        End If
     
        If mintVersion >= 100 Then
            Set objItem = lvwPara.ListItems.Add(, "sga_target", "sga_target")
            objItem.SubItems(2) = ">0"
            objItem.SubItems(3) = "数据缓存和共享池等共享内存的总量，0表示自动管理，建议分配为总的物理内存*80%*80%，如果修改值大于SGA_MAX_SIZE，则需先修改后者并重启实例"
            objItem.SubItems(4) = "数字"
            objItem.SubItems(5) = "转换强调"
            strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        End If
        
        '11G及以上
        If mintVersion > 100 Then
            Set objItem = lvwPara.ListItems.Add(, "memory_target", "memory_target")
            objItem.SubItems(2) = ">0"
            objItem.SubItems(3) = "数据库实例的内存总量，0表示自动管理，建议设置为物理内存的80%，如果修改值大于memory_max_target，则需先修改后者并重启实例。"
            objItem.SubItems(4) = "数字"
            objItem.SubItems(5) = "转换强调"
            strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        End If
    End With
    
    Dim strParList As String
    strParList = Replace(strParas, "'", "")
    Dim lng系数 As Long, bln不符要求 As Boolean
    With rsTemp
        On Error Resume Next
        strSql = "select lower(name) as name,value" & _
                " from v$parameter" & _
                " where name in (" & Mid(strParas, 2) & ")"
        If .State = adStateOpen Then .Close
        .Open strSql, gcnOracle, adOpenKeyset
        Do While Not .EOF
            bln不符要求 = False
            Set objItem = lvwPara.ListItems(.Fields("name").value)
            objItem.SubItems(1) = .Fields("value").value
            Select Case objItem.SubItems(4)
            Case "数字"
                lng系数 = 1
                If objItem.SubItems(5) = "转换强调" Then lng系数 = 1048576  '(1024 * 1024)
                objItem.SubItems(1) = Format(Val(.Fields("value").value) / lng系数, "0") & IIf(lng系数 > 1, "M", "")
                            
                If .Fields("value").value < Val(Mid(objItem.SubItems(2), 3)) * lng系数 Then
                    If objItem.Key = "sga_target" Then
                        blnSGA = True
                    End If
                    If objItem.SubItems(5) = "强调" Or objItem.SubItems(5) = "转换强调" Then
                        objItem.ForeColor = vbBlue
                        objItem.ListSubItems(1).ForeColor = vbBlue
                        objItem.ListSubItems(2).ForeColor = vbBlue
                        objItem.ListSubItems(3).ForeColor = vbBlue
                    End If
                End If
            Case "文本"
                If objItem.Key = "optimizer_mode" Then
                    blnCBO = .Fields("value").value = "ALL_ROWS"
                End If
                
                If UCase(.Fields("value").value) <> objItem.SubItems(2) Then
                    If objItem.SubItems(5) = "强调" Then
                        objItem.ForeColor = vbBlue
                        objItem.ListSubItems(1).ForeColor = vbBlue
                        objItem.ListSubItems(2).ForeColor = vbBlue
                        objItem.ListSubItems(3).ForeColor = vbBlue
                    Else
                        bln不符要求 = True
                    End If
                End If
            Case "兼容版本号"
                If Val(Replace(.Fields("value").value, ".", "")) < Val(Replace(Mid(objItem.SubItems(2), 3), ".", "")) Then
                    objItem.ForeColor = vbBlue
                    objItem.ListSubItems(1).ForeColor = vbBlue
                    objItem.ListSubItems(2).ForeColor = vbBlue
                    objItem.ListSubItems(3).ForeColor = vbBlue
                End If
            End Select
            If bln不符要求 Then
                '不合要求(红)
                objItem.ForeColor = RGB(255, 0, 0)
                objItem.ListSubItems(1).ForeColor = RGB(255, 0, 0)
                objItem.ListSubItems(2).ForeColor = RGB(255, 0, 0)
                objItem.ListSubItems(3).ForeColor = RGB(255, 0, 0)
            End If
            .MoveNext
        Loop
        
        
        '-- 清除不显示的项
        Dim i As Integer
 
        For i = 1 To lvwPara.ListItems.Count
            If blnCBO = False Then
                If i < lvwPara.ListItems.Count Then
                    If InStr("optimizer_index_cost_adj,optimizer_index_caching", lvwPara.ListItems(i).Key) > 0 Then
                         lvwPara.ListItems.Remove i
                         i = i - 1
                    End If
                End If
            End If
            
            If blnSGA = True Then
                If i < lvwPara.ListItems.Count Then
                    If InStr("db_cache_size,shared_pool_size,java_pool_size", lvwPara.ListItems(i).Key) > 0 Then
                        lvwPara.ListItems.Remove i
                        i = i - 1
                    End If
                End If
            End If
        Next

        
    End With
    
End Sub

Private Sub txtMem_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) <= 0 Then KeyAscii = 0
End Sub


