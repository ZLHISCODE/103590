VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAppUpgradePreproccess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "升迁前置检查"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11655
   Icon            =   "frmAppUpgradePreproccess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11655
      TabIndex        =   1
      Top             =   0
      Width           =   11655
      Begin VB.Frame fraTop 
         Height          =   120
         Left            =   0
         TabIndex        =   5
         Top             =   840
         Width           =   11880
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   $"frmAppUpgradePreproccess.frx":6852
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   11400
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11655
      TabIndex        =   0
      Top             =   6555
      Width           =   11655
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出(&E)"
         Height          =   350
         Left            =   10440
         TabIndex        =   8
         Top             =   360
         Width           =   1100
      End
      Begin VB.CommandButton cmdRecheck 
         Caption         =   "重新检查(&R)"
         Height          =   350
         Left            =   7600
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdjust 
         Caption         =   "调整(&A)"
         Height          =   350
         Left            =   9000
         TabIndex        =   6
         Top             =   360
         Width           =   1100
      End
      Begin VB.Frame fraBottom 
         Height          =   120
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   11880
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsCheckResult 
      Height          =   5820
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   11460
      _cx             =   20214
      _cy             =   10266
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   0
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483628
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   100
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAppUpgradePreproccess.frx":6908
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   4
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList imgEdit 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":6A20
               Key             =   "Check"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":6FBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":7554
               Key             =   "签名"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":78A6
               Key             =   "Woman"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":E108
               Key             =   "Man"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":1496A
               Key             =   "UnCheck"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":14E32
               Key             =   "AllCheck"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAppUpgradePreproccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum SysCheck
    SC_检查对象 = 0
    SC_当前值 = 1
    SC_建议值 = 2
    SC_调整 = 3
    SC_检查说明 = 4
    SC_处理警告 = 5
End Enum

Private mrsCheckInfo        As ADODB.Recordset
Private mstrUsers           As String
Private mbln10G             As Boolean
Private Const SQL_CAPTION = "升迁前置检查"
Private mlngBeach           As Long
'去掉SYS与SYSTEM
Private Const mstrOracleUser      As String = "'ANONYMOUS','AURORA$JIS$UTILITY$','AURORA$ORB$UNAUTHENTICATED','CTXSYS','DBSNMP','DIP','DMSYS','DVF','DVSYS','EXFSYS','HR','LBACSYS','MDDATA','MDSYS','MGMT_VIEW','OAS_PUBLIC','ODM','ODM_MTR','OE','OGG','OLAPSYS','ORDPLUGINS','ORDSYS','OSE$HTTP$ADMIN','OUTLN','PERFSTAT','PM','QS','QS_ADM','QS_CB','QS_CBADM','QS_CS','QS_ES','QS_OS','QS_WS','REPADMIN','RMAN','SCOTT','SH','SI_INFORMTN_SCHEMA','SYSMAN','TRACESVR','TSMSYS','WEBSYS','WKPROXY','WKSYS','WKUSER','WK_TEST','WMSYS','XDB'"
Private mblnOK              As Boolean
Private mblnExecBefore      As Boolean
Private Enum CheckClass
    CC_SYSPARA = 0
    CC_DBFile = 1
    CC_AutoJob = 2
    CC_Trigger = 3
    CC_Scheduler = 4
    CC_Privs = 5
End Enum

Private Enum ChceckType
    CT_CheckAndLoad = 0
    CT_OnlyCheck = 1
    CT_OnlyLoad = 2
End Enum
'******************************************************************************************************************
'功能：检查系统状态能否进行升迁
'blnExecBefore-是否提前升级
'返回：TRUE-可以进行升迁，false-不能进行升迁
'******************************************************************************************************************
Public Function ShowMe(ByVal blnExecBefore As Boolean) As Boolean
    mblnExecBefore = blnExecBefore
    If mblnExecBefore Then ShowMe = True: Exit Function
    mstrUsers = GetUsers
    mbln10G = GetOracleVersion(True, True) < 11
    mblnOK = False
    If Not gblnDBA Then
        MsgBox """" & gstrUserName & """不是DBA，无法进行升级前检查，如要进行升级前检查，请授予""" & gstrUserName & """DBA权限。", vbInformation, gstrSysName
        ShowMe = True
        Exit Function
    End If
    If Not LoadAllCheck(CT_OnlyCheck) Then
        ShowMe = True
        Exit Function
    End If
    '检查存在问题，弹出界面提示用户修正
    If mrsCheckInfo.RecordCount <> 0 Then
        Me.Show vbModal, frmMDIMain
        ShowMe = mblnOK
    Else
        ShowMe = True
    End If
End Function
'******************************************************************************************************************
'功能：检查并加载检查结果
'参数：intType=0，检查并加载，1-只检查，不加在，2-只加载不检查
'******************************************************************************************************************
Private Function LoadAllCheck(Optional ByVal intType As Integer) As Boolean
    On Error GoTo errH
    If intType < 2 Then
        mlngBeach = 0
        Set mrsCheckInfo = CopyNewRec(Nothing, True, , Array("分类序号", adInteger, Empty, Empty, "检查分类", adVarChar, 100, Empty, _
                            "对象序号", adInteger, Empty, Empty, "检查对象", adVarChar, 100, Empty, "所有者", adVarChar, 100, Empty, "对象", adVarChar, 100, Empty, _
                            "当前值", adVarChar, 100, Empty, "建议值", adVarChar, 100, Empty, _
                            "修正SQL", adVarChar, 100, Empty, "调整类型", adInteger, Empty, Empty, "是否调整", adInteger, Empty, Empty, "是否DBA", adInteger, Empty, Empty, _
                            "检查说明", adVarChar, 200, Empty, "处理警告", adVarChar, 200, Empty, "是否已修正", adInteger, Empty, Empty))
        '调整类型 0-不可调整且不允许不调整，1-不可用调整单允许不调整，2-可调整
        Call ShowFlash("正在检查系统参数，请稍候！")
        If Not CheckSysPara Then Call ShowFlash(""): Exit Function
        Call ShowFlash("正在检查数据库文件，请稍候！")
        If Not CheckDBFile Then Call ShowFlash(""): Exit Function
        Call ShowFlash("正在检查自动作业和系统调度，请稍候！")
        If Not CheckAutoJobs Then Call ShowFlash(""): Exit Function
        Call ShowFlash("正在检查触发器，请稍候！")
        If Not CheckTriggers Then Call ShowFlash(""): Exit Function
        Call ShowFlash("正在对象权限，请稍候！")
        If Not CheckPrivs Then Call ShowFlash(""): Exit Function
        Call ShowFlash("")
    End If
    If intType <> 1 Then
        Call LoadCheckInfo
    End If
    LoadAllCheck = True
    Exit Function
errH:
    Call ShowFlash("")
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function


'******************************************************************************************************************
'功能：获取ZLHIS系统的所有者
'******************************************************************************************************************
Private Function GetUsers() As String
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select f_List2str(Cast(Collect(Chr(39) || 所有者 || Chr(39)) As t_Strlist)) 所有者" & vbNewLine & _
            "From (Select Upper(所有者) 所有者" & vbNewLine & _
            "       From zlBakSpaces" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select Upper(所有者)" & vbNewLine & _
            "       From zlSystems" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select 'ZLTOOLS'" & vbNewLine & _
            "       From Dual)"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    GetUsers = rsTmp!所有者 & ""
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "获取标准系统用户出现错误：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'功能：检查项目添加到记录集
'参数：strOwner=问题对象的所有者
'      strName=问题对象名称
'      strItem=问题对象展示名称
'      strCurValue=问题对象当前状态
'      lngCheckClass=问题对象的分类ID
'      strSuggestiveValue=问题对象的建议状态
'      strAdjustSQL=问题对象的修正SQL,不存在时，不能自动修复
'      strCheckInfo=问题对象的检查说明
'      strAdjustWarn=问题对象的修正警告，部分对象需要特殊处理或手工处理，均在此处
'      blnIgnor=问题对象的是否可以忽略
'      blnDBA=问题对象的修正是否需要DBA身份，没有修正SQL的对象，请注意传NULL
'******************************************************************************************************************
Private Sub AddCheckItem(ByVal strOwner As String, strName As String, ByVal strItem As String, ByVal strCurValue As String, ByVal lngCheckClass As Long, ByVal strCheckClass As String, _
                        ByVal strSuggestiveValue As String, ByVal strAdjustSQL As String, ByVal strCheckInfo As String, _
                        ByVal strAdjustWarn As String, Optional ByVal blnIgnor As Boolean, Optional ByVal blnDBA As Boolean)
     mrsCheckInfo.AddNew Array("分类序号", "检查分类", "对象序号", "检查对象", "所有者", "对象", "当前值", "建议值", "修正SQL", "调整类型", "是否DBA", "检查说明", "处理警告", "是否调整", "是否已修正"), _
                        Array(lngCheckClass, strCheckClass, mrsCheckInfo.RecordCount + 1, strItem, strOwner, strName, strCurValue, strSuggestiveValue, IIf(strAdjustSQL = "", Null, strAdjustSQL), IIf(blnIgnor, IIf(strAdjustSQL = "", 1, 2), 0), IIf(blnDBA, 1, 0), strCheckInfo, IIf(strAdjustSQL <> "" And strAdjustWarn = "", "自动调整", strAdjustWarn), IIf(strAdjustSQL = "", 0, 1), 0)

End Sub
'******************************************************************************************************************
'功能：检查数据库参数
'******************************************************************************************************************
Private Function CheckSysPara() As Boolean
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select Name , Value From V$parameter Where Name =[1] And Value =[2]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, "optimizer_index_cost_adj", "100")
    If Not rsTmp.EOF Then Call AddCheckItem("", rsTmp!name & "", rsTmp!name & "", rsTmp!value & "", CC_SYSPARA, "数据库参数", "20", "alter system set " & rsTmp!name & "=20", "缺省值100会导致产品性能问题", "", True, True)
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, "optimizer_index_caching", "0")
    If Not rsTmp.EOF Then Call AddCheckItem("", rsTmp!name & "", rsTmp!name & "", rsTmp!value & "", CC_SYSPARA, "数据库参数", "80", "alter system set " & rsTmp!name & "=80", "缺省0会导致产品性能问题", "", True, True)
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, "O7_DICTIONARY_ACCESSIBILITY", "FALSE")
    If Not rsTmp.EOF Then Call AddCheckItem("", rsTmp!name & "", rsTmp!name & "", rsTmp!value & "", CC_SYSPARA, "数据库参数", "TRUE", "", "导致系统视图无法授权，影响升级以及产品功能", "请手工调整为TRUE后重启数据库", False, False)

    strSQL = "Select Name , Value From V$parameter Where Name = [1] And Zl_To_Number(Value) < [2]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, "log_buffer", "104857600")
    If Not rsTmp.EOF Then Call AddCheckItem("", rsTmp!name & "", rsTmp!name & "", Int(Val(rsTmp!value & "") / 1024 / 1024) & "M", CC_SYSPARA, "数据库参数", ">=100M", "", "影响系统升级中数据修正效率", "请手工调整为至少100M后重启数据库", True, False)
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, "parallel_execution_message_size", "8192")
    If Not rsTmp.EOF Then Call AddCheckItem("", rsTmp!name & "", rsTmp!name & "", rsTmp!value & "", CC_SYSPARA, "数据库参数", ">=8192", "", "影响系统升级并行执行", "请手工调整为8192后重启数据库", True, False)
    
    CheckSysPara = True
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "检查数据库参数出现错误：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'功能：检查日志文件
'******************************************************************************************************************
Private Function CheckDBFile() As Boolean
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    Dim strFile As String
    
    On Error GoTo errH
    strSQL = "Select 'INST_ID:' || a.Inst_Id || ',GROUP:' || a.Group# Name,b.Member," & vbNewLine & _
            "       a.Bytes Value" & vbNewLine & _
            "From Gv$log A" & vbNewLine & _
            "Join Gv$logfile B" & vbNewLine & _
            "On (a.Group# = b.Group# And a.Inst_Id = b.Inst_Id)" & vbNewLine & _
            "Where a.Bytes < 104857600" & vbNewLine & _
            "Order By a.Inst_Id, a.Group#, a.Thread#, b.Member"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    Do While Not rsTmp.EOF
        strFile = GetFileNameByPath(rsTmp!Member & "")
        Call AddCheckItem("", rsTmp!name & "," & strFile, rsTmp!name & "," & strFile, Int(Val(rsTmp!value & "") / 1024 / 1024) & "M", CC_DBFile, "数据库文件", ">=100M", "", "影响系统升级中数据修正效率", "请手工调整为至少100M", True, False)
        rsTmp.MoveNext
    Loop
    CheckDBFile = True
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "检查日志文件出现错误：" & err.Description, vbInformation, gstrSysName
    CheckDBFile = True
    err.Clear
End Function
'******************************************************************************************************************
'功能：检查自动作业
'******************************************************************************************************************
Private Function CheckAutoJobs() As Boolean
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    Dim strProcName     As String
    
    On Error GoTo errH
    '检查执行时间在当前时间下限2小时，上限5小时，仍旧未禁止的自动作业
    strSQL = "Select a.Job, a.Broken, a.Schema_User, Upper(What) What" & vbNewLine & _
            "From Dba_Jobs A" & vbNewLine & _
            "Where a.Job In (Select 作业号 From Zltools.Zlautojobs) And a.Broken = 'N' And" & vbNewLine & _
            "      Nvl(a.Next_Date, Sysdate + 10) Between Sysdate - 2 / 24 And Sysdate + 5 / 24"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    
    Do While Not rsTmp.EOF
        Call AddCheckItem(rsTmp!Schema_User & "", rsTmp!Job & "", rsTmp!Job & "(" & rsTmp!What & ")", rsTmp!Broken, CC_AutoJob, "自动作业", "BROKEN=Y", "Dbms_Job.Broken(" & rsTmp!Job & ", True)", "影响相关表的数据修正脚本执行效率", "升级期间禁用", True, rsTmp!Schema_User & "" = "SYS" Or rsTmp!Schema_User & "" = "SYSTEM")
        rsTmp.MoveNext
    Loop
    '非ZLHIS管理的自动作业
    strSQL = "Select a.Job, a.Broken, a.Schema_User, Upper(What) What" & vbNewLine & _
            "From Dba_Jobs A" & vbNewLine & _
            "Where a.Schema_User Not In (" & mstrOracleUser & ") And a.Job Not In (Select 作业号 From Zltools.Zlautojobs) And" & vbNewLine & _
            "      a.Broken = 'N' And Nvl(a.Next_Date, Sysdate + 10) Between Sysdate - 2 / 24 And Sysdate + 5 / 24"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    
    Do While Not rsTmp.EOF
        strProcName = GetJobProcedure(rsTmp!What & "")
        If strProcName <> "" Then
            If ObjectReferencedZLHIS(rsTmp!Schema_User, strProcName) Then
                Call AddCheckItem(rsTmp!Schema_User & "", rsTmp!Job & "", rsTmp!Job & "(" & rsTmp!What & ")", rsTmp!Broken, CC_AutoJob, "自动作业", "BROKEN=Y", "Dbms_Job.Broken(" & rsTmp!Job & ", True)", "影响相关表的数据修正脚本执行效率", "升级期间禁用", True, rsTmp!Schema_User & "" = "SYS" Or rsTmp!Schema_User & "" = "SYSTEM")
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    CheckAutoJobs = True
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "检查自动作业出现错误：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'功能：检查发器
'******************************************************************************************************************
Private Function CheckTriggers() As Boolean
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    
    On Error GoTo errH
    'ZLHIS所有的触发器不用进一步判断对象
    If CheckAndAdjustMustTable("ZLTABLES") Then
        strSQL = "Select a.Owner, a.Trigger_Name, a.Status" & vbNewLine & _
                "From Dba_Triggers A" & vbNewLine & _
                "Where a.Table_Name In (Select b.表名 From Zltables B Where b.分类 Not Like 'A%') And a.Status = 'ENABLED' And a.Trigger_Type <> 'INSTEAD OF' And" & vbNewLine & _
                "      a.Table_Owner In (" & mstrUsers & ")"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    Else
        strSQL = "Select a.Owner, a.Trigger_Name,a.Status From Dba_Triggers A Where a.Status = 'ENABLED' And a.Table_Owner In (" & mstrUsers & ") And a.Trigger_Type <> 'INSTEAD OF'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    End If
    Do While Not rsTmp.EOF
        Call AddCheckItem(rsTmp!Owner & "", rsTmp!trigger_name & "", rsTmp!Owner & "." & rsTmp!trigger_name, "ENABLED", CC_Trigger, "触发器", "DISABLED", "alter trigger " & rsTmp!Owner & "." & rsTmp!trigger_name & " disable", "影响该表的数据修正脚本执行效率", "升级期间禁用", True, rsTmp!Owner & "" = "SYS" Or rsTmp!Owner & "" = "SYSTEM")
        rsTmp.MoveNext
    Loop
    
    CheckTriggers = True
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "检查触发器出现错误：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'功能：检查特殊用户的对象权限
'******************************************************************************************************************
Private Function CheckPrivs() As Boolean
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    
    On Error GoTo errH
    '升级是使用该表进行数据修正
    '检查ZLTOOLS与PUBLIC权限
    strSQL = "Select a.Grantee, a.Owner, a.Table_Name, a.Privilege" & vbNewLine & _
            "From (Select 'ZLTOOLS' Grantee, 'SYS' Owner, 'DBA_ROLE_PRIVS' Table_Name, 'SELECT' Privilege From Dual) A" & vbNewLine & _
            "Where Not Exists (Select 1" & vbNewLine & _
            "       From Dba_Tab_Privs C" & vbNewLine & _
            "       Where c.Owner = 'SYS' And (c.Grantee = 'PUBLIC' Or a.Grantee<>'PUBLIC' And c.Grantee = 'ZLTOOLS') And" & vbNewLine & _
            "             c.Table_Name = a.Table_Name And c.Privilege = a.Privilege)"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    Do While Not rsTmp.EOF
        Call AddCheckItem(rsTmp!Owner & "", rsTmp!Table_Name & "", rsTmp!Grantee & " " & rsTmp!Privilege & " On " & rsTmp!Owner & "." & rsTmp!Table_Name, "", CC_Privs, "对象权限", rsTmp!Privilege & "", "Grant " & rsTmp!Privilege & " On " & rsTmp!Owner & "." & rsTmp!Table_Name, "升级可能会出现异常，以及影响产品使用", "", False, rsTmp!Owner & "" = "SYS")
        rsTmp.MoveNext
    Loop
    CheckPrivs = True
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "检查对象权限出现错误：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'功能：检查一个对象是否引用了ZLHIS系统的基础对象(表),表为存储数据的基本对象，引用该对象的上级对象可能会导致死锁，因此只检查表
'参数：strOwner-对象所有者
'      strObjectName=对象名称
'      strObjectType=对象类型，当不传时，默认不为同义词
'返回：TRUE-引用了ZLHIS的对象，false-未引用ZLHIS对象
'******************************************************************************************************************
Private Function ObjectReferencedZLHIS(ByVal strOwner As String, ByVal strObjectName As String, Optional ByVal strObjectType As String) As Boolean
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select Count(1) 计数" & vbNewLine & _
            "From (Select a.Owner, a.Name, a.Type, a.Referenced_Owner, a.Referenced_Name, a.Referenced_Type" & vbNewLine & _
            "       From All_Dependencies A" & vbNewLine & _
            "       Start With a.Owner = [1] And a.Name = [2] And a.Type " & IIf(strObjectType = "", "<>'SYNONYM'", "= [3]") & vbNewLine & _
            "       Connect By Prior a.Referenced_Owner = a.Owner And Prior a.Referenced_Name = a.Name And" & vbNewLine & _
            "                  Prior a.Referenced_Type = a.Type) B" & vbNewLine & _
            "Where b.Referenced_Owner In (" & mstrUsers & ") And b.Referenced_Type = 'TABLE' Or b.Owner In (" & mstrUsers & ") And b.Type = 'TABLE'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, strOwner, strObjectName, strObjectType)
    ObjectReferencedZLHIS = rsTmp!计数 <> 0
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "检查对象依赖出现错误：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'功能：解析JOB执行内容，获取执行的存储过程
'参数：strWhat-JOB执行内容
'返回：JOB执行的存储过程
'******************************************************************************************************************
Private Function GetJobProcedure(ByVal strWhat As String) As String
    Dim arrTmp          As Variant
    Dim strProcedure    As String
    '测试用例EMD_MAINTENANCE.EXECUTE_EM_DBMS_JOB_PROCS();
    arrTmp = Split(strWhat & ";", ";")
    strProcedure = arrTmp(0)
    arrTmp = Split(strProcedure & "(", "(")
    strProcedure = arrTmp(0)
    arrTmp = Split(strProcedure & ".", ".") '.分割放在(之后，因为可能传参，传参中存在.
    strProcedure = arrTmp(0)
    GetJobProcedure = strProcedure
End Function
'******************************************************************************************************************
'功能：根据文件路径获取文件名
'参数：strFilePath-文件路径，可能是Linux系统下的路径
'返回：文件名称
'******************************************************************************************************************
Private Function GetFileNameByPath(ByVal strFilePath As String) As String
    Dim lngPos  As Long
    
    lngPos = InStrRev(strFilePath, "/")
    If lngPos = 0 Then
        lngPos = InStrRev(strFilePath, "\")

    End If
    If lngPos = 0 Then
        GetFileNameByPath = strFilePath
    Else
        GetFileNameByPath = Mid(strFilePath, lngPos + 1)
    End If
End Function


'******************************************************************************************************************
'功能：将已经处理的系统调度、自动作业、触发器记录到数据库
'******************************************************************************************************************
Private Sub AdjustRecordToDB()
    Dim blnAutoJobs     As Boolean
    Dim strScheduler    As String
    Dim strSQL          As String
    Dim strJobs         As String
    
    On Error GoTo errH
    With mrsCheckInfo
        '将处理的触发器记录到数据库
        .Filter = "分类序号=" & CC_Trigger & " And 是否已修正=" & mlngBeach
        .Sort = "分类序号,对象序号"
        If Not .EOF Then
            Call SetUpgradeConfig("触发器状态", "0")
            Do While Not .EOF
                strSQL = strSQL & " Union All Select '" & !所有者 & "' 所有者,'" & !对象 & "' 名称 From Dual"
                .MoveNext
            Loop
            strSQL = Mid(strSQL, Len(" Union All ") + 1)
            strSQL = "Insert Into Zltriggers (所有者, 名称) Select 所有者, 名称" & vbNewLine & _
                    "From (" & strSQL & ") A" & vbNewLine & _
                    "Where Not Exists (Select 1 From Zltriggers B Where b.名称 = a.名称 And b.所有者 = a.所有者)"
            gcnOracle.Execute strSQL, , adCmdText
        End If
        .Filter = "分类序号=" & CC_AutoJob & " And 是否已修正=" & mlngBeach
        .Sort = "分类序号,对象序号"
        blnAutoJobs = False
        If Not .EOF Then
            blnAutoJobs = True
            Do While Not .EOF
                If Not SetAutoJobs(Val(!对象 & "")) Then
                    strJobs = strJobs & "," & Val(!对象 & "")
                End If
                .MoveNext
            Loop
            If strJobs <> "" Then
                Call SetUpgradeConfig("禁用的自动作业", Mid(strJobs, 2))
            End If
        End If
        '将处理的系统调度以及自动作业记录到数据库
        .Filter = "分类序号=" & CC_Scheduler & " And 是否已修正=" & mlngBeach
        .Sort = "分类序号,对象序号"
        If Not .EOF Then
            blnAutoJobs = True
            Do While Not .EOF
                strScheduler = strScheduler & ",'" & !对象 & "'"
                
                .MoveNext
            Loop
            Call SetUpgradeConfig("禁用的系统调度", Mid(strScheduler, 2))
        End If
        If blnAutoJobs Then Call SetUpgradeConfig("后台作业状态", "0")
    End With
    Exit Sub
errH:
    MsgBox "自动修正出现错误：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

'******************************************************************************************************************
'功能：将系统调度、自动作业、触发器配置记录到数据库
'参数:strConfigName-配置名称
'     strConfigValue=配置值
'******************************************************************************************************************
Private Sub SetUpgradeConfig(ByVal strConfigName As String, ByVal strConfigValue As String)
    Dim strSQL      As String
    Dim lngAffect   As Long
    Dim strTmp      As String, rsTmp    As ADODB.Recordset
    
    On Error GoTo errH
    If strConfigName = "禁用的系统调度" Or strConfigName = "禁用的自动作业" Then
        strSQL = "Select 内容 From Zlupgradeconfig Where 项目 = '" & strConfigName & "'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        If rsTmp.EOF Then
            strSQL = "Insert Into ZLTOOLS.zlUpgradeConfig(项目,内容) values('" & strConfigName & "'," & strConfigValue & ")"
            gcnOracle.Execute strSQL, lngAffect, adCmdText
        Else
            If rsTmp!内容 & "" = "" Then
                strTmp = strConfigValue
            Else
                strTmp = rsTmp!内容 & "," & strConfigValue
            End If
            strSQL = "Update Zlupgradeconfig Set 内容 = " & strConfigValue & " Where 项目 = '" & strConfigName & "'"
            gcnOracle.Execute strSQL, lngAffect, adCmdText
        End If
    Else
        strSQL = "Update Zlupgradeconfig Set 内容 = " & strConfigValue & " Where 项目 = '" & strConfigName & "'"
        gcnOracle.Execute strSQL, lngAffect, adCmdText
        If lngAffect = 0 Then
            strSQL = "Insert Into ZLTOOLS.zlUpgradeConfig(项目,内容) values('" & strConfigName & "'," & strConfigValue & ")"
            gcnOracle.Execute strSQL, lngAffect, adCmdText
        End If
    End If
    Exit Sub
errH:
    MsgBox "自动修正出现错误：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

'******************************************************************************************************************
'功能：标记停用的后台自动作业
'参数:strJobNum=自动作业号
'返回：True-存在这样的自动作业，FALSE-不存在这样的自动作业
'******************************************************************************************************************
Private Function SetAutoJobs(ByVal strJobNum As String) As Boolean
    Dim strSQL      As String
    Dim lngAffect   As Long
    
    On Error GoTo errH
    strSQL = "Update zlAutoJobs Set 系统升级停用 = 1 Where 作业号 = " & strJobNum
    gcnOracle.Execute strSQL, lngAffect, adCmdText
    SetAutoJobs = lngAffect <> 0
    Exit Function
errH:
    MsgBox "自动修正出现错误：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function

'******************************************************************************************************************
'功能：将检查结果加载到界面
'******************************************************************************************************************
Private Sub LoadCheckInfo()
    Dim lngRow          As Long, lngLastClassRow    As Long
    Dim lngLastClass    As Long, blnHideClass       As Boolean
    Dim blnSelALl       As Boolean, lngCanCheck     As Long, lngAllCanCheck As Long
    With vsCheckResult
        .Redraw = False
        .OutlineCol = 0
        mrsCheckInfo.Filter = ""
        mrsCheckInfo.Sort = "分类序号,对象序号"
        .Rows = vsCheckResult.FixedRows
        lngLastClass = -1
        lngCanCheck = 0
        Do While Not mrsCheckInfo.EOF
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            If lngLastClass <> mrsCheckInfo!分类序号 Then
                If lngLastClass <> -1 Then
                    .RowHidden(lngLastClassRow) = blnHideClass
                    If lngCanCheck > 0 Then
                        Set .Cell(flexcpPicture, lngLastClassRow, SC_调整) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
                        .Cell(flexcpData, lngLastClassRow, SC_调整) = IIf(blnSelALl, 1, 0)
                    Else
                        .Cell(flexcpData, lngLastClassRow, SC_调整) = 0
                    End If
                End If
                blnHideClass = True
                blnSelALl = True
                
                .Cell(flexcpData, lngRow, SC_检查对象) = Val(mrsCheckInfo!分类序号 & "")
                .TextMatrix(lngRow, SC_检查对象) = mrsCheckInfo!检查分类
                .IsSubtotal(lngRow) = True
                .RowOutlineLevel(lngRow) = 1
                .RowData(lngRow) = 0
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H8000000F
                
                lngLastClassRow = lngRow
                lngLastClass = mrsCheckInfo!分类序号
                lngCanCheck = 0
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
            .TextMatrix(lngRow, SC_检查对象) = mrsCheckInfo!检查对象
            .TextMatrix(lngRow, SC_当前值) = mrsCheckInfo!当前值
            .Cell(flexcpData, lngRow, SC_调整) = Val(mrsCheckInfo!调整类型 & "")
            If mrsCheckInfo!调整类型 = 0 Then
                .Cell(flexcpForeColor, lngRow, SC_检查对象, lngRow, SC_处理警告) = &HFF0000
            End If
            If Val(mrsCheckInfo!调整类型 & "") = 2 Then
                .Cell(flexcpChecked, lngRow, SC_调整, lngRow, SC_调整) = IIf(mrsCheckInfo!是否调整 = 0, flexUnchecked, flexChecked)
                If mrsCheckInfo!是否调整 = 0 Then blnSelALl = False
                If mrsCheckInfo!是否已修正 = 0 Then
                    lngCanCheck = lngCanCheck + 1
                    lngAllCanCheck = lngAllCanCheck + 1
                End If
            Else
                .Cell(flexcpChecked, lngRow, SC_调整, lngRow, SC_调整) = 0
            End If
            .TextMatrix(lngRow, SC_建议值) = mrsCheckInfo!建议值
            .TextMatrix(lngRow, SC_检查说明) = mrsCheckInfo!检查说明
            .TextMatrix(lngRow, SC_处理警告) = mrsCheckInfo!处理警告
            
            .RowOutlineLevel(lngRow) = 2
            .IsSubtotal(lngRow) = True
            .RowData(lngRow) = Val(mrsCheckInfo!对象序号 & "")
            .RowHidden(lngRow) = mrsCheckInfo!是否已修正 > 0
            If Not .RowHidden(lngRow) Then blnHideClass = False

            mrsCheckInfo.MoveNext
        Loop
        
        If lngLastClass <> -1 Then
            If lngCanCheck > 0 Then
                Set .Cell(flexcpPicture, lngLastClassRow, SC_调整) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
                .Cell(flexcpData, lngLastClassRow, SC_调整) = IIf(blnSelALl, 1, 0)
            Else
                .Cell(flexcpData, lngLastClassRow, SC_调整) = 0
            End If
            .Cell(flexcpPictureAlignment, .FixedRows, SC_调整, .Rows - 1, SC_调整) = flexAlignCenterCenter
            .RowHidden(lngLastClassRow) = blnHideClass
        End If
        cmdAdjust.Enabled = lngAllCanCheck > 0
        If lngAllCanCheck <> 0 Then
            Set .Cell(flexcpPicture, 0, SC_调整) = imgEdit.ListImages("AllCheck").Picture
            .ColData(SC_调整) = 1
        End If
        .Redraw = True
    End With
End Sub

Private Sub cmdAdjust_Click()
    Dim cnDBA       As ADODB.Connection
    Dim cnTmp       As ADODB.Connection
    Dim blnNewBeach As Boolean
    mrsCheckInfo.Filter = "是否DBA=1 And 是否调整=1 And 修正SQL<>NULL And 是否已修正=0"
    If mrsCheckInfo.RecordCount <> 0 Then
        Set cnDBA = GetConnection("SYSTEM")
        If cnDBA Is Nothing Then Exit Sub
    End If
    With mrsCheckInfo
        .Filter = "是否调整=1 And 修正SQL<>NULL And 是否已修正=0"
        .Sort = "分类序号,对象序号"
        Do While Not .EOF
            If !是否DBA = 0 Then
                Set cnTmp = gcnOracle
            Else
                Set cnTmp = cnDBA
            End If
            If ExecuteCmdText(!修正SQL, Me.Caption, cnTmp, !分类序号 = CC_AutoJob, True) = "" Then
                If Not blnNewBeach Then
                    mlngBeach = mlngBeach + 1
                    blnNewBeach = True
                End If
                .Update "是否已修正", mlngBeach
            End If
            .MoveNext
        Loop
    End With
    If blnNewBeach Then Call AdjustRecordToDB
    mrsCheckInfo.Filter = "是否调整=1 And 是否已修正=0"
    If mrsCheckInfo.RecordCount = 0 Then
        mblnOK = True
        Unload Me
        Exit Sub
    End If
    Call LoadAllCheck(CT_OnlyLoad)
End Sub

Public Function ExecuteCmdText(ByVal strSQL As String, ByVal strFormCaption As String, Optional cnOracle As ADODB.Connection, Optional ByVal blnProcedure As Boolean, Optional ByVal blnErrResume As Boolean) As String
'功能：执行无返回值语句
    If blnErrResume Then
        On Error Resume Next
    Else
        On Error GoTo errH
    End If
    If blnProcedure Then
        cnOracle.Execute strSQL, , adCmdStoredProc
    Else
        cnOracle.Execute strSQL, , adCmdText
    End If
    If err.Number <> 0 Then
        ExecuteCmdText = err.Description
        err.Clear
    End If
    Exit Function
errH:
    ExecuteCmdText = err.Description
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Private Sub cmdExit_Click()
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdRecheck_Click()
    Call LoadAllCheck(CT_CheckAndLoad)
End Sub

Private Sub Form_Load()
    Call LoadAllCheck(CT_OnlyLoad)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If MsgBox("请完成调整或确认升迁预处理，否则将无法进行升级。是否退出？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsCheckInfo = Nothing
End Sub

Private Sub vsCheckResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim blnSelALl   As Boolean
    Dim lngClassRow As Long, i      As Long
    
    With vsCheckResult
        blnSelALl = True
        If Col <> SC_调整 Then Exit Sub
        Call RecUpdate(mrsCheckInfo, "对象序号=" & Val(.RowData(Row)), "是否调整", IIf(.Cell(flexcpChecked, Row, SC_调整, Row, SC_调整) = flexUnchecked, 0, 1))
        For i = Row + 1 To .Rows - 1
            If .RowData(i) = 0 Then
                Exit For
            Else
                If .Cell(flexcpData, i, SC_调整) = 2 Then
                    If .Cell(flexcpChecked, i, SC_调整, i, SC_调整) = flexUnchecked Then
                        blnSelALl = False
                    End If
                End If
            End If
        Next
        
        For i = Row To .FixedRows Step -1
            If .RowData(i) = 0 Then
                lngClassRow = i
                Exit For
            Else
                If .Cell(flexcpData, i, SC_调整) = 2 Then
                    If .Cell(flexcpChecked, i, SC_调整, i, SC_调整) = flexUnchecked Then
                        blnSelALl = False
                    End If
                End If
            End If
        Next
        If Not .Cell(flexcpPicture, lngClassRow, SC_调整) Is Nothing Then
            Set .Cell(flexcpPicture, lngClassRow, SC_调整) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
            .Cell(flexcpData, lngClassRow, SC_调整) = IIf(blnSelALl, 1, 0)
        End If
    End With
End Sub

Private Sub vsCheckResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > -1 And NewCol > -1 Then
        If vsCheckResult.RowData(NewRow) = 0 Then
            vsCheckResult.BackColorSel = &H8000000F
        Else
            vsCheckResult.BackColorSel = &H8000000D
        End If
    End If
End Sub

Private Sub vsCheckResult_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = SC_调整
End Sub

Private Sub vsCheckResult_Click()
    Dim blnSelALl       As Boolean
    Dim i               As Long
    With vsCheckResult
        If .MouseRow > 0 And .MouseCol = SC_调整 Then
            If .RowData(.Row) = 0 Then
                blnSelALl = .Cell(flexcpData, .MouseRow, SC_调整) = 0
                Call RecUpdate(mrsCheckInfo, "分类序号=" & .Cell(flexcpData, .MouseRow, SC_检查对象) & " And 是否已修正=0", "是否调整", IIf(blnSelALl, 1, 0))
                For i = .Row + 1 To .Rows - 1
                    If .RowData(i) = 0 Then
                        Exit For
                    Else
                        If Not .RowHidden(i) Then
                            If .Cell(flexcpData, i, SC_调整) = 2 Then
                                .Cell(flexcpChecked, i, SC_调整, i, SC_调整) = IIf(blnSelALl, flexChecked, flexUnchecked)
                            End If
                        End If
                    End If
                Next
                If Not .Cell(flexcpPicture, .MouseRow, SC_调整) Is Nothing Then
                    Set .Cell(flexcpPicture, .MouseRow, SC_调整) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
                    .Cell(flexcpData, .MouseRow, SC_调整) = IIf(blnSelALl, 1, 0)
                End If
            End If
        ElseIf .MouseRow = 0 And .Col = SC_调整 Then
            If Not .Cell(flexcpPicture, 0, SC_调整) Is Nothing Then
                blnSelALl = Val(.ColData(SC_调整)) = 0
                Call RecUpdate(mrsCheckInfo, "", "是否调整", IIf(blnSelALl, 1, 0))
                For i = .FixedRows To .Rows - 1
                    If .RowData(i) = 0 Then
                        If Not .Cell(flexcpPicture, i, SC_调整) Is Nothing Then
                            Set .Cell(flexcpPicture, i, SC_调整) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
                            .Cell(flexcpData, i, SC_调整) = IIf(blnSelALl, 1, 0)
                        End If
                    Else
                        If .Cell(flexcpData, i, SC_调整) = 2 Then
                            .Cell(flexcpChecked, i, SC_调整, i, SC_调整) = IIf(blnSelALl, flexChecked, flexUnchecked)
                        End If
                    End If
                Next
                .Cell(flexcpPicture, 0, SC_调整) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
                .ColData(SC_调整) = IIf(blnSelALl, 1, 0)
            End If
        End If
    End With
End Sub

Private Sub vsCheckResult_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> SC_调整 Then
        Cancel = True
    Else
        If vsCheckResult.RowData(Row) = 0 Then
            Cancel = True
        End If
    End If
End Sub
