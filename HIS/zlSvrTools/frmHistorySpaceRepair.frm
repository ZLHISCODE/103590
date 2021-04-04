VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmHistorySpaceRepair 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "历史库结构修复"
   ClientHeight    =   6780
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "frmHistorySpaceRepair.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdRepair 
      Caption         =   "修复(&R)"
      Height          =   350
      Left            =   8280
      TabIndex        =   5
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   9480
      TabIndex        =   4
      Top             =   6000
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   6408
      Width           =   10704
      _ExtentX        =   18891
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmHistorySpaceRepair.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15319
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "16:31"
            Key             =   "STANUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraCheck 
      Height          =   5985
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   10680
      Begin VSFlex8Ctl.VSFlexGrid vsCheckResult 
         Height          =   5100
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   10395
         _cx             =   18336
         _cy             =   8996
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
         BackColorSel    =   16764057
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   100
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHistorySpaceRepair.frx":0E1C
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
         OutlineBar      =   0
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
         WordWrap        =   -1  'True
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
      End
      Begin VB.Frame fraTop 
         Height          =   120
         Left            =   15
         TabIndex        =   2
         Top             =   570
         Width           =   10680
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "根据转出表的定义，检查在线库与历史库的结构一致性"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   3
         Top             =   225
         Width           =   5400
      End
   End
   Begin ComctlLib.ImageList ist 
      Left            =   120
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHistorySpaceRepair.frx":0F57
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHistorySpaceRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSys  As Long                   '当前系统的系统编号
Private mstrVersion   As String            '当前系统的版本号
Private mstrBakOwnerName As String         '当前系统在线非只读的历史表空间的所有者
Private mstrOwnerName As String

Private mcnBakDB As New ADODB.Connection '历史表空间所有者建立的连接

Private marrBakAddSQL() As Variant '以后备库管理员登录执行的SQL，晚于在线库管理员先登录执行的SQL的执行
Private marrOnlineAddSQL() As Variant '在线库管理员先登录执行的SQL

Private mblnSucced As Boolean '修复是否成功
Private mblnUpdate As Boolean '是否是升级程序调用
Private mblnFirstAct As Boolean '窗体是否首次激活
Private mblnAllRepair As Boolean '是否修复成功
Private mblnCurDB As Boolean '是否是当前系统的当前历史库
Private mrsErrInfo As ADODB.Recordset

Private mrsSQL As ADODB.Recordset
Private mlngIndex As Long '记录SQL顺序
Private mstrBakDB           As String   '历史表空间
Private mstrBakIndexDB      As String   '历史库索引表空间
Private mstrBakLobDB        As String   '历史库LOB表空间
Private mstrDBLink  As String   '带有@符号
Private mstrServer As String

Private Enum RepCols
    RC_DifInfo = 0
    RC_DifType = 1
    RC_TabName = 2
    RC_ObjName = 3
    RC_ColName = 4
    RC_ObjType = 5
    RC_ObjLen = 6
    RC_ObjScale = 7
    RC_AutoRep = 8
    RC_RepSQL = 9
    RC_RepMethod = 10
    
End Enum

Private Enum DifType
    DT_HLackTab = 0 '历史表缺失
    DT_HMoreCol = 1 '历史表多一列
    DT_HLessCol = 2 '历史表少一列
    DT_HDataTypeDif = 3 '列数类型不同
    DT_HRepLenDif = 4 '可修复列长度或精度差异
    DT_HNotRepLenDif = 5 '不可修复的列长度或精度差异
    DT_HLobTablespace = 6 'LOB字段表的表空间错误
    DT_HIndUsable = 7 '历史表失效的索引
    DT_HIndDel = 8 '历史表多余的索引
    DT_HIndAdd = 9 '历史表缺少的索引
    DT_HIndColDif = 10 '历史表索引列差异
    DT_HConDisable = 11 '历史表禁用的约束
    DT_URefConDel = 12 '子表数据未转出
    DT_HConDel = 13 '历史表多余的约束
    DT_HConAdd = 14 '历史表缺少的约束
    DT_HConColDIf = 15 '历史表约束列差异
    DT_HIndexTablesapce = 16 '索引表空间错误
End Enum

Public Function ShowRepair(ByVal frmMain As Form, ByVal lng系统 As Long, ByVal blnUpdate As Boolean, Optional ByVal strBakUser As String, Optional ByVal strBakDB As String, Optional ByVal blnCurDB As Boolean = True, Optional ByRef rsRepairSQL As ADODB.Recordset, Optional cnDBBAK As ADODB.Connection, Optional strDbLink As String) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '功能:修复历史表空间的数据结构
    '参数:cnOracle-系统连接
    '     strOwner-所有者用户名
    '     lng系统-系统号
    '     blnUpdate -升级时的结构修复
    '     strDBLink=DBLInk名称
    '返回:安装成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    mlngSys = lng系统

    '系统权限控制
    strSQL = "select 所有者,版本号,名称 from zlSystems where 编号=" & mlngSys
    Call OpenRecordset(rsTemp, strSQL, "读取所有者")
    
    If Not rsTemp.EOF Then
        mstrOwnerName = Nvl(rsTemp!所有者)
        mstrVersion = Nvl(rsTemp!版本号)
    Else
        If Not blnUpdate Then MsgBox "系统不存在,可能被他人拆卸,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If gstrUserName <> mstrOwnerName Then
        If Not blnUpdate Then MsgBox "你不是当前应用程序的所有者,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If mstrVersion <> "" Then
        If Val(Split(mstrVersion, ".")(0)) < 10 Then
                If Not blnUpdate Then MsgBox "不支持9以下的版本,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
        End If
    End If
    mblnCurDB = False
    mstrServer = gstrServer
    If Not blnUpdate Then
        '历史表空间所有者，以及表空间名称
        strSQL = "Select 名称,所有者,DB连接 From Zltools.Zlbakspaces Where 系统 = " & mlngSys & "  And 当前 = 1 And 只读 = 0"
        Call OpenRecordset(rsTemp, strSQL, "读取历史表空间所有者")
        If rsTemp.EOF Then
            MsgBox "当前没有可用的历史数据空间或者历史数据空间目前的状态为只读,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        Else
            mstrBakOwnerName = Nvl(rsTemp!所有者)
            mblnCurDB = True
            mstrDBLink = Nvl(rsTemp!DB连接)
            mstrBakDB = Nvl(rsTemp!名称)
        End If
    Else
        mstrBakOwnerName = strBakUser
        mblnCurDB = blnCurDB
        mstrDBLink = strDbLink
        mstrBakDB = strBakDB
    End If
    mstrBakDB = UCase(mstrBakDB)
    If mstrDBLink <> "" Then
        strSQL = "Select Owner, Db_Link, Username, Host" & vbNewLine & _
                    "From All_Db_Links" & vbNewLine & _
                    "Where Owner =[1] And Username =[2] And Db_Link||'.' Like [3]"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取DBLink服务器", gstrUserName, UCase(mstrBakOwnerName), UCase(mstrDBLink) & ".%")
        If Not rsTemp.EOF Then mstrServer = rsTemp!Host & ""
    End If
    mstrDBLink = IIf(mstrDBLink = "", "", "@") & mstrDBLink
    '获取索引表空间与LOB表空间
    strSQL = "Select a.Name From V$tablespace" & mstrDBLink & " a Where a.Name Like '" & mstrBakDB & "_%'"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取历史库表空间")
    rsTemp.Filter = "Name='" & mstrBakDB & "_IDX'"
    If Not rsTemp.EOF Then
        mstrBakIndexDB = rsTemp!name
    Else
        mstrBakIndexDB = mstrBakDB
    End If
    rsTemp.Filter = "Name='" & mstrBakDB & "_LOB'"
    If Not rsTemp.EOF Then
        mstrBakLobDB = rsTemp!name
    Else
        mstrBakLobDB = mstrBakDB
    End If
    mblnFirstAct = True
    mblnAllRepair = True
    mblnUpdate = blnUpdate
    If blnUpdate Then
        Set mcnBakDB = cnDBBAK
        Set mrsSQL = rsRepairSQL
        If mrsSQL Is Nothing Then
            Set mrsSQL = GetIniRec
        End If
        Call LoadCheckData
    End If
    On Error Resume Next
    If Not blnUpdate Then
        Me.Show 1
        On Error GoTo 0
        ShowRepair = mblnSucced
    Else
        Set rsRepairSQL = mrsSQL
        Exit Function
    End If
    On Error GoTo 0
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRepair_Click()
    Dim i As Long
    Dim comTmp As New ADODB.Command
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngCount As Long

    '历史表空间所有者验证
    If Not frmUserCheckLogin.ShowLogin(UCT_CurZLBAK, mcnBakDB, mstrBakOwnerName, mstrServer, mlngSys) Then Exit Sub
    

    lngCount = 5
    
    mblnAllRepair = False
    
    Call SetFaceCtlEnable
    If mcnBakDB Is Nothing Then Exit Sub
    
    On Error Resume Next
    SetPromptText ("（1/" & lngCount & ")正在将在线库中相关表授权给历史库用户")
    comTmp.CommandType = adCmdText
    Set comTmp.ActiveConnection = gcnOracle
    For i = LBound(marrOnlineAddSQL) To UBound(marrOnlineAddSQL)
        comTmp.CommandText = marrOnlineAddSQL(i)
        comTmp.Execute
        If err <> 0 Then
            Call AddErrIntoRs(0, err.Description, , , marrOnlineAddSQL(i))
            err.Clear
        End If
    Next
    
    SetPromptText ("（2/" & lngCount & ")开始修复历史库数据结构")
    With vsCheckResult
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, RC_RepSQL) <> "" Then
                Select Case Val(.TextMatrix(i, RC_DifType))
                    Case DT_URefConDel
                        Set comTmp.ActiveConnection = gcnOracle
                    Case DT_HConAdd
                        '检查索引是否存在，存在，则删除
                        strSQL = "Select /*+rule*/" & vbNewLine & _
                                    " 1" & vbNewLine & _
                                    "From User_Indexes A" & vbNewLine & _
                                    "Where A.Index_Name = '" & .TextMatrix(i, RC_ObjName) & "'"
                         Call OpenRecordset(rsTmp, strSQL, "数据转移相关过程有效检查", , , mcnBakDB)
                         Set comTmp.ActiveConnection = mcnBakDB
                         If Not rsTmp.EOF Then
                            comTmp.CommandText = " Drop Index  " & .TextMatrix(i, RC_ObjName)
                            comTmp.Execute
                            If err <> 0 Then
                                Call AddErrIntoRs(IIf(Val(.TextMatrix(i, RC_DifType)) = DT_URefConDel, 0, 1), err.Description, .TextMatrix(i, RC_TabName), .TextMatrix(i, RC_ObjName), " Drop Index  " & .TextMatrix(i, RC_ObjName))
                                err.Clear
                            End If
                         End If
                    Case Else
                        Set comTmp.ActiveConnection = mcnBakDB
                End Select
                
                comTmp.CommandText = .TextMatrix(i, RC_RepSQL)
                comTmp.Execute
                If err <> 0 Then
                    Call AddErrIntoRs(IIf(Val(.TextMatrix(i, RC_DifType)) = DT_URefConDel, 0, 1), err.Description, .TextMatrix(i, RC_TabName), .TextMatrix(i, RC_ObjName), .TextMatrix(i, RC_RepSQL))
                    err.Clear
                End If
            End If
        Next
    End With
    SetPromptText ("（3/" & lngCount & ")历史库授权在线库所有者")
    Set comTmp.ActiveConnection = mcnBakDB
    For i = LBound(marrBakAddSQL) To UBound(marrBakAddSQL)
        comTmp.CommandText = marrBakAddSQL(i)
        comTmp.Execute
        If err <> 0 Then
            Call AddErrIntoRs(0, err.Description, , , marrBakAddSQL(i))
            err.Clear
        End If
    Next
    
    SetPromptText ("（4/" & lngCount & ")开始重新创建历史库H视图并授权")
    If mstrDBLink = "" Then
        Set comTmp.ActiveConnection = gcnOracle
        Call GrantBakToUser(mcnBakDB, mstrOwnerName)
    End If
    If mblnCurDB Then
        Call CreateAppView(mstrOwnerName, mstrBakOwnerName, mlngSys, mstrDBLink)
    End If
    SetPromptText ("（5/" & lngCount & ")数据转移相关过程有效检查与重编译")
    Set comTmp.ActiveConnection = gcnOracle
    strSQL = "Select 'Alter Procedure Zl" & mlngSys \ 100 & "_Datamove_Tag compile' As Sql" & vbNewLine & _
            "From User_Objects A" & vbNewLine & _
            "Where a.Object_Name = Upper('Zl" & mlngSys \ 100 & "_Datamove_Tag') And a.Object_Type = 'PROCEDURE' And a.Status = 'INVALID'" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select 'Alter Procedure ' || a.Object_Name || ' compile' As Sql" & vbNewLine & _
            "From User_Objects A" & vbNewLine & _
            "Where a.Object_Name In (Upper('Zl" & mlngSys \ 100 & "_Datamoveout1'), Upper('Zl_Retu_Clinic'), Upper('Zl_Retu_Exes')) And" & vbNewLine & _
            "      a.Object_Type = 'PROCEDURE' And a.Status = 'INVALID'"
    Call OpenRecordset(rsTmp, strSQL, "数据转移相关过程有效检查")
    While Not rsTmp.EOF
        comTmp.CommandText = rsTmp!SQL & ""
        comTmp.Execute
        If err <> 0 Then
            Call AddErrIntoRs(0, err.Description, , , strSQL)
            err.Clear
        End If
        rsTmp.MoveNext
    Wend
    
    SetPromptText ("修复完成")
    Call LoadErrInfo(mrsErrInfo)
    mblnAllRepair = True
    Call SetFaceCtlEnable
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnFirstAct Then
        mblnFirstAct = False
        Me.Refresh
        If Not LoadCheckData Then
            SetPromptText ("检查完成,历史库未发现结构问题")
        Else
            SetPromptText ("检查完成,请修复")
        End If
        If mrsErrInfo Is Nothing Then Set mrsErrInfo = New ADODB.Recordset
        With mrsErrInfo
            .Fields.Append "数据库", adInteger
            .Fields.Append "差异类型", adInteger
            .Fields.Append "差异信息", adVarChar, 100
            .Fields.Append "表名", adVarChar, 50
            .Fields.Append "对象名", adVarChar, 50
            .Fields.Append "错误SQL", adVarChar, 200
            .Open
        End With
    End If
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    marrBakAddSQL = Array()
    marrOnlineAddSQL = Array()
    Call ApplyOEM(stbThis)
End Sub

Private Function LoadCheckData() As Boolean

    '----------------------------------------------------------------------------------------------------------------------------------
    '功能:加载历史表空间数据结构检查结果
    '返回:检查成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strLackTable    As String
    Dim strLobTable     As String
    Dim lngTotal        As Long
    Dim lngCur          As Long
    
    If Not mblnUpdate Then
        vsCheckResult.Redraw = False
    End If
    On Error GoTo errH:
    lngTotal = 15
    '（一、） 历史表缺失
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表")
    strSQL = "Select '历史表缺失' 差异信息, " & DT_HLackTab & " 差异类型, t.表名, Null 对象名, Null 列名, Null 数据类型, Null 长度, Null 精度, '√' 自动修复, '创建此表' 修复说明" & vbNewLine & _
            "From Zltools.Zlbaktables t, (Select Table_Name From All_Tables" & mstrDBLink & " a Where a.Owner = '" & UCase(mstrBakOwnerName) & "') b" & vbNewLine & _
            "Where t.系统 = " & mlngSys & " And b.Table_Name(+) = t.表名 And b.Table_Name Is Null"
    
    Call OpenRecordset(rsTmp, strSQL, "历史库缺失表")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp, True)
    End If
    If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
    While Not rsTmp.EOF
        strLackTable = strLackTable & ",'" & UCase(rsTmp!表名 & "") & "'"
        rsTmp.MoveNext
    Wend
    If Len(strLackTable) <> 0 Then strLackTable = Mid(strLackTable, 2)
    '(二、) 列检查： 1-历史表缺少列，无需修复 2-历史表有多余列，添加该列) 该类型排除了“表的缺失”
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & "/" & lngTotal & ")正在检查历史表表列")
    '(1)现查找历史表空间存在的表
    strSQL = IIf(Len(strLackTable) = 0, "", " And T.表名 Not In (" & strLackTable & ")")
    '(2)查找历史表空间，缺少或多余的列
    strSQL = "Select  Decode(a.Column_Name, Null, '历史表有多余列', '历史表缺少列') 差异信息, Decode(B.Column_Name, Null, " & DT_HLessCol & ", " & DT_HMoreCol & ") 差异类型," & vbNewLine & _
            "       Nvl(b.Table_Name, a.Table_Name) 表名, Nvl(b.Column_Name, a.Column_Name) 对象名, Nvl(b.Column_Name, a.Column_Name) 列名," & vbNewLine & _
            "       Nvl(b.Data_Type, a.Data_Type) 数据类型," & vbNewLine & _
            "       Decode(Nvl(b.Data_Type, a.Data_Type),  'XMLTYPE',Null,'DATE',Null,'Long Raw',Null,'BLOB',Null,'CLOB',Null," & vbNewLine & _
            "       Nvl(Nvl(b.Data_Precision, a.Data_Precision),Nvl(b.Data_Length, a.Data_Length))) 长度, Nvl(b.Data_Scale, a.Data_Scale) 精度," & vbNewLine & _
            "       Decode(a.Column_Name, Null, '×', '√') 自动修复," & vbNewLine & _
            "       Decode(a.Column_Name, Null, '无需修复', '新增此列') 修复说明" & vbNewLine & _
            "From (Select c.Table_Name,c.Column_Name,c.Data_Type,c.Data_Precision,c.Data_Scale,c.Data_Length From User_Tab_Columns C, Zltools.Zlbaktables T Where c.Table_Name = t.表名 And t.系统 = " & mlngSys & strSQL & ") A" & vbNewLine & _
            "Full Join (Select d.Table_Name,d.Column_Name,d.Data_Type,d.Data_Precision,d.Data_Scale,d.Data_Length " & vbNewLine & _
            "           From All_Tab_Columns" & mstrDBLink & " D, Zltools.Zlbaktables T" & vbNewLine & _
            "           Where d.Table_Name = t.表名 And t.系统 = " & mlngSys & strSQL & "  And d.Owner = '" & UCase(mstrBakOwnerName) & "') B" & vbNewLine & _
            "On a.Table_Name = b.Table_Name And a.Column_Name = b.Column_Name" & vbNewLine & _
            "Where a.Column_Name Is Null Or b.Column_Name Is Null"

    Call OpenRecordset(rsTmp, strSQL, "历史库列检查")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    
    '(三、）列类型差异
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表表列类型")
    strSQL = "Select '列类型差异' 差异信息, " & DT_HDataTypeDif & " 差异类型, b.Table_Name 表名, b.Column_Name 对象名, b.Column_Name 列名, b.Data_Type 数据类型," & vbNewLine & _
                    "       Decode(b.Data_Type,'XMLTYPE',Null,'DATE',Null,'Long Raw',Null,'BLOB',Null,'CLOB',Null,Nvl(b.Data_Precision, b.Data_Length)) 长度," & vbNewLine & _
                    "        '×' 自动修复, '修复说明待定' 修复说明" & vbNewLine & _
                    "From User_Tab_Columns a, All_Tab_Columns" & mstrDBLink & " b, (Select t.表名 From Zltools.Zlbaktables t Where t.系统 = " & mlngSys & ") c" & vbNewLine & _
                    "Where a.Table_Name = b.Table_Name And b.Table_Name = c.表名 And b.Owner = '" & UCase(mstrBakOwnerName) & "' And a.Column_Name = b.Column_Name And" & vbNewLine & _
                    "      a.Data_Type <> b.Data_Type "
    Call OpenRecordset(rsTmp, strSQL, "列类型不同的列")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '(四、) 列的精度以及长度差异(4-可修复，5-不可修复） 差异类型等信息经过程序判断得出
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表表列精度")
    strSQL = "Select Null 差异信息, -1 差异类型, 表名, 对象名, 对象名 列名, 数据类型, a.后备长度, a.后备精度, a.在线长度, a.在线精度, Null 自动修复, Null 修复说明" & vbNewLine & _
            "From (Select a.Table_Name 表名, a.Column_Name 对象名, a.Data_Type 数据类型," & vbNewLine & _
            "              Decode(a.Data_Type, 'XMLTYPE',Null,'DATE',Null,'Long Raw',Null, 'BLOB',Null,'CLOB',Null, Nvl(a.Data_Precision,a.Data_Length)) 在线长度, a.Data_Scale 在线精度," & vbNewLine & _
            "              Decode(b.Data_Type, 'XMLTYPE',Null,'DATE',Null,'Long Raw',Null, 'BLOB',Null,'CLOB',Null,Nvl(b.Data_Precision,b.Data_Length)) 后备长度, b.Data_Scale 后备精度" & vbNewLine & _
            "       From User_Tab_Columns A, All_Tab_Columns" & mstrDBLink & " B" & vbNewLine & _
            "       Where a.Table_Name = b.Table_Name And b.Owner = '" & UCase(mstrBakOwnerName) & "' And a.Column_Name = b.Column_Name And Exists" & vbNewLine & _
            "        (Select 1 From Zltools.Zlbaktables T Where t.系统 = " & mlngSys & " And t.表名 = a.Table_Name) And a.Data_Type = b.Data_Type) A" & vbNewLine & _
            "Where 在线长度 <> 后备长度 Or 在线精度 <> 后备精度"


    Call OpenRecordset(rsTmp, strSQL, "列的精度差异")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    
    '(S)历史表LOB表空间检查，必须放在索引约束之前检查，因为修正可能会导致索引约束失效
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表LOB表空间")
    strSQL = "Select '历史库表表空间错误' 差异信息, " & DT_HLobTablespace & " 差异类型, a.Table_Name 表名, Null 对象名, Null 列名, Null 数据类型, Null 长度, Null 精度, '×' 自动修复," & vbNewLine & _
                    "       '建议手工将该表移动到表空间" & mstrBakLobDB & "' 修复说明" & vbNewLine & _
                    "From All_Tables" & mstrDBLink & " a" & vbNewLine & _
                    "Where a.Owner = '" & UCase(mstrBakOwnerName) & "'" & vbNewLine & _
                    "And a.Table_Name In (Select Distinct c.Table_Name" & vbNewLine & _
                    "                    From User_Tab_Cols c, Zltools.Zlbaktables t" & vbNewLine & _
                    "                    Where c.Table_Name = t.表名" & vbNewLine & _
                    "                    And t.系统 = " & mlngSys & vbNewLine & _
                    "                    And c.Data_Type In ('BLOB', 'CLOB', 'BFILE', 'XMLTYPE'))" & vbNewLine & _
                    "And a.Tablespace_Name Not in( '" & mstrBakLobDB & "'" & IIf(mstrBakLobDB = mstrBakDB, ",'" & mstrBakLobDB & "_LOB')", ")")
    Call OpenRecordset(rsTmp, strSQL, "历史表约束列差异")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
    While Not rsTmp.EOF
        strLobTable = strLobTable & ",'" & UCase(rsTmp!表名 & "") & "'"
        rsTmp.MoveNext
    Wend
    strLobTable = Mid(strLobTable, 2)
    '(六）修复索引
    '（1）非主/唯一键，非块数据索引的有效性检查
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表索引的有效性")
    strSQL = "Select '历史表中失效索引' 差异信息, " & DT_HIndUsable & "  差异类型, a.Table_Name 表名, a.Index_Name 对象名, a.Colstr 列名, Null 数据类型, Null 长度, Null 精度, '√' 自动修复," & vbNewLine & _
                    "       '索引重建' 修复说明" & vbNewLine & _
                    "From (Select d.Table_Name, d.Index_Name," & vbNewLine & _
                    "              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Ind_Columns d, Zltools.Zlbaktables t" & vbNewLine & _
                    "       Where d.Table_Name = t.表名 And t.系统 = " & mlngSys & " And Instr(d.Index_Name, '_PK') = 0 And Instr(d.Index_Name, '_UQ') = 0 And Instr(d.Index_Name, '_IX_待转出') = 0" & vbNewLine & _
                    "       Group By d.Table_Name, d.Index_Name) a," & vbNewLine & _
                    "     (Select e.Table_Name, f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Cons_Columns e, User_Constraints f, Zltools.Zlbaktables t, User_Constraints c," & vbNewLine & _
                    "            (Select 表名 From Zltools.Zlbaktables" & vbNewLine & _
                    "              Union All Select '病案主页' From Dual" & vbNewLine & _
                    "              Union All Select '病人信息' From Dual) g" & vbNewLine & _
                    "       Where e.Table_Name = t.表名 And t.系统 = " & mlngSys & " And e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And" & vbNewLine & _
                    "             c.Constraint_Name = f.r_Constraint_Name And g.表名(+) = c.Table_Name And g.表名 Is Null" & vbNewLine & _
                    "       Group By e.Table_Name, e.Constraint_Name) b," & vbNewLine & _
                    "     (Select Index_Name" & vbNewLine & _
                    "       From All_Indexes" & mstrDBLink & " k, Zltools.Zlbaktables t" & vbNewLine & _
                    "       Where k.Table_Owner = '" & UCase(mstrBakOwnerName) & "' And k.Status = 'UNUSABLE' And k.Index_Type <> 'LOB' And k.Table_Name = t.表名 And" & vbNewLine & _
                    "             t.系统 = " & mlngSys & ") h" & vbNewLine & _
                    "Where a.Table_Name = b.Table_Name(+) And a.Colstr = b.Colstr(+) And b.Colstr Is Null And a.Index_Name = h.Index_Name"

    Call OpenRecordset(rsTmp, strSQL, "历史表中失效索引")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    
    '（2）历史表空间外键上的索引的删除（若外键引用的表非历史表空间存在的表，则需删除)
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表多余的索引")
    
    strSQL = "Select '历史表多余的索引' 差异信息, " & DT_HIndDel & " 差异类型, a.Table_Name 表名, a.Index_Name 对象名, a.Colstr 列名, Null 数据类型, Null 长度, Null 精度, '√' 自动修复," & vbNewLine & _
                "       '删除索引' 修复说明" & vbNewLine & _
                "From (Select d.Table_Name, d.Index_Name," & vbNewLine & _
                "              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
                "       From User_Ind_Columns d, Zltools.Zlbaktables t" & vbNewLine & _
                "       Where d.Table_Name = t.表名 And t.系统 = " & mlngSys & " And Instr(d.Index_Name, '_PK') = 0 And Instr(d.Index_Name, '_UQ') = 0" & vbNewLine & _
                "       Group By d.Table_Name, d.Index_Name) a," & vbNewLine & _
                "     (Select e.Table_Name, f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr" & vbNewLine & _
                "       From User_Cons_Columns e, User_Constraints f, Zltools.Zlbaktables t, User_Constraints c," & vbNewLine & _
                "            (Select 表名 From Zltools.Zlbaktables" & vbNewLine & _
                "              Union All Select '病案主页' From Dual" & vbNewLine & _
                "              Union All Select '病人信息' From Dual) g" & vbNewLine & _
                "       Where e.Table_Name = t.表名 And t.系统 = " & mlngSys & " And e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And" & vbNewLine & _
                "             c.Constraint_Name = f.r_Constraint_Name And g.表名(+) = c.Table_Name And g.表名 Is Null" & vbNewLine & _
                "       Group By e.Table_Name, e.Constraint_Name) b," & vbNewLine & _
                "     (Select Index_Name" & vbNewLine & _
                "       From All_Indexes" & mstrDBLink & " k, Zltools.Zlbaktables t" & vbNewLine & _
                "       Where k.Table_Owner = '" & UCase(mstrBakOwnerName) & "'  And k.Table_Name = t.表名 And t.系统 = " & mlngSys & ") h" & vbNewLine & _
                "Where a.Table_Name = b.Table_Name And a.Colstr = b.Colstr And h.Index_Name = a.Index_Name" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select '历史表多余的索引' 差异信息, " & DT_HIndDel & " 差异类型, k.Table_Name 表名, k.Index_Name 对象名, '待转出' 列名, Null 数据类型, Null 长度, Null 精度, '√' 自动修复," & vbNewLine & _
                "       '删除索引' 修复说明" & vbNewLine & _
                "From All_Indexes" & mstrDBLink & " k, Zltools.Zlbaktables t" & vbNewLine & _
                "Where k.Table_Owner = '" & UCase(mstrBakOwnerName) & "'  And k.Table_Name = t.表名 And t.系统 = " & mlngSys & " And k.Index_Name Like '%_待转出'"
    Call OpenRecordset(rsTmp, strSQL, "历史表多余的索引")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '(3)历史表空间外键上的索引的添加（若外键引用的表是历史表空间存在的表，则需添加)
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表缺少的索引")
    strSQL = "Select  '历史表缺少的索引' 差异信息, " & DT_HIndAdd & " 差异类型, a.Table_Name 表名, a.Index_Name 对象名, a.Colstr 列名, Null 数据类型, Null 长度, Null 精度, '√' 自动修复," & vbNewLine & _
                    "       '创建索引' 修复说明" & vbNewLine & _
                    "From (Select d.Table_Name, d.Index_Name,f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Ind_Columns D, Zltools.Zlbaktables T" & vbNewLine & _
                    "       Where d.Table_Name = t.表名 And t.系统 = " & mlngSys & " And Instr(d.Index_Name, '_PK') = 0 And Instr(d.Index_Name, '_UQ') = 0 And Instr(d.Index_Name, '_IX_待转出')=0" & vbNewLine & _
                    "       Group By d.Table_Name, d.Index_Name) A," & vbNewLine & _
                    "     (Select e.Table_Name, f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Cons_Columns e, User_Constraints f, Zltools.Zlbaktables t, User_Constraints c," & vbNewLine & _
                    "            (Select 表名 From Zltools.Zlbaktables" & vbNewLine & _
                    "              Union All Select '病案主页' From Dual" & vbNewLine & _
                    "              Union All Select '病人信息' From Dual) g" & vbNewLine & _
                    "       Where e.Table_Name = t.表名 And t.系统 = " & mlngSys & " And e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And" & vbNewLine & _
                    "             c.Constraint_Name = f.r_Constraint_Name And g.表名(+) = c.Table_Name And g.表名 Is Null" & vbNewLine & _
                    "       Group By e.Table_Name, e.Constraint_Name) b," & vbNewLine & _
                    "     (Select Index_Name" & vbNewLine & _
                    "       From All_Indexes" & mstrDBLink & " k, Zltools.Zlbaktables t" & vbNewLine & _
                    "       Where k.Table_Owner = '" & UCase(mstrBakOwnerName) & "' And k.Table_Name = t.表名 And t.系统 = " & mlngSys & ") h" & vbNewLine & _
                    "Where a.Table_Name = b.Table_Name(+) And a.Colstr = b.Colstr(+) And b.Colstr Is Null and a.Index_Name=h.index_name(+) and h.index_name is null"
    Call OpenRecordset(rsTmp, strSQL, "历史表缺少的索引")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '(4)历史表空间索引列发生变化
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表索引列的差异")
    strSQL = "Select '索引列差异' 差异信息, " & DT_HIndColDif & " 差异类型, u.Table_Name 表名, u.Index_Name 对象名, u.Colstr 列名, Null 数据类型, Null 长度, Null 精度, '√' 自动修复," & vbNewLine & _
                    "       '重建索引' 修复说明" & vbNewLine & _
                    "From (Select a.Table_Name, a.Index_Name, f_List2str(Cast(Collect(a.Column_Name Order By a.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Ind_Columns A, Zltools.Zlbaktables T" & vbNewLine & _
                    "       Where a.Table_Name = t.表名 And t.系统 = " & mlngSys & " And Instr(a.Index_Name, '_PK') = 0 And Instr(a.Index_Name, '_UQ') = 0 And Instr(a.Index_Name, '_IX_待转出')=0" & vbNewLine & _
                    "       Group By a.Table_Name, a.Index_Name) U," & vbNewLine & _
                    "     (Select a.Table_Name, a.Index_Name, f_List2str(Cast(Collect(a.Column_Name Order By a.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From All_Ind_Columns" & mstrDBLink & " A, Zltools.Zlbaktables T" & vbNewLine & _
                    "       Where a.Table_Name = t.表名 And t.系统 = " & mlngSys & " And a.Table_Owner ='" & UCase(mstrBakOwnerName) & "'" & vbNewLine & _
                    "       Group By a.Table_Name, a.Index_Name) H," & vbNewLine & _
                    "     (Select e.Table_Name, f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Cons_Columns e, User_Constraints f, Zltools.Zlbaktables t, User_Constraints c," & vbNewLine & _
                    "            (Select 表名 From Zltools.Zlbaktables" & vbNewLine & _
                    "              Union All Select '病案主页' From Dual" & vbNewLine & _
                    "              Union All Select '病人信息' From Dual) g" & vbNewLine & _
                    "       Where e.Table_Name = t.表名 And t.系统 = " & mlngSys & " And e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And" & vbNewLine & _
                    "             c.Constraint_Name = f.r_Constraint_Name And g.表名(+) = c.Table_Name And g.表名 Is Null" & vbNewLine & _
                    "       Group By e.Table_Name, e.Constraint_Name) b" & vbNewLine & _
                    "Where u.Table_Name = h.Table_Name And u.Index_Name = h.Index_Name And u.Colstr <> h.Colstr And" & vbNewLine & _
                    "      u.Table_Name = b.Table_Name(+) And u.Colstr = b.Colstr(+) And b.Table_Name Is Null"
    Call OpenRecordset(rsTmp, strSQL, "索引列差异")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '（五、）修复约束
    '（1）主/唯一键约束的有效性检查
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史库中禁用的约束")
    strSQL = "Select '历史表中禁用的约束' 差异信息, " & DT_HConDisable & " 差异类型, a.Table_Name 表名, a.Constraint_Name 对象名," & vbNewLine & _
            "       f_List2str(Cast(Collect(b.Column_Name Order By b.Position) As t_Strlist)) 列名, a.Constraint_Type 数据类型, Null 长度, Null 精度, '√' 自动修复," & vbNewLine & _
            "       '启用约束' 修复说明" & vbNewLine & _
            "From All_Constraints" & mstrDBLink & " A, All_Cons_Columns" & mstrDBLink & " B, Zltools.Zlbaktables T" & vbNewLine & _
            "Where a.Owner = '" & UCase(mstrBakOwnerName) & "'  And a.Owner = b.Owner And a.Constraint_Type In ('P', 'U') And a.Table_Name = t.表名 And t.系统 = " & mlngSys & " And" & vbNewLine & _
            "      a.Status = 'DISABLED' And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From User_Constraints C" & vbNewLine & _
            "       Where c.Constraint_Name = a.Constraint_Name And c.Constraint_Type = a.Constraint_Type) And a.Owner = b.Owner And" & vbNewLine & _
            "      a.Constraint_Name = b.Constraint_Name" & vbNewLine & _
            "Group By a.Table_Name, a.Constraint_Name, a.Constraint_Type"

    Call OpenRecordset(rsTmp, strSQL, "约束有效性")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '(2)在线库需删除的外键约束
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查在线库转出表的外键约束")
    strSQL = "Select  Distinct '子表数据未转出' 差异信息, " & DT_URefConDel & " 差异类型, d.Table_Name 表名, d.Constraint_Name 对象名," & vbNewLine & _
            "                f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) 列名, Null 数据类型, Null 长度, Null 精度, '×' 自动修复," & vbNewLine & _
            "                '不修复' 修复说明" & vbNewLine & _
            "From (Select Table_Name, Constraint_Name, Owner" & vbNewLine & _
            "       From (Select a.Owner, a.r_Constraint_Name, a.Constraint_Name, a.Table_Name, b.Table_Name r_Table_Name" & vbNewLine & _
            "              From User_Constraints A, User_Constraints B" & vbNewLine & _
            "              Where a.r_Constraint_Name = b.Constraint_Name(+)) C" & vbNewLine & _
            "       Start With c.r_Table_Name In (Select t.表名 From Zltools.Zlbaktables T Where t.系统 = " & mlngSys & ")" & vbNewLine & _
            "       Connect By Nocycle Prior c.Table_Name = c.r_Table_Name) D, User_Cons_Columns E" & vbNewLine & _
            "Where Not Exists (Select 1 From Zltools.Zlbaktables T Where t.表名 = d.Table_Name) And" & vbNewLine & _
            "      e.Constraint_Name = d.Constraint_Name And e.Table_Name = d.Table_Name And e.Owner = d.Owner" & vbNewLine & _
            "Group By d.Constraint_Name, d.Table_Name"
            
    Call OpenRecordset(rsTmp, strSQL, "子表数据未转出")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    
    '(3)历史表空间需删除的约束（对于非主键与唯一键的约束均需删除,历史表空间中的主键唯一键没有对应在线数据库的也许删除）
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表多余约束")
    strSQL = "Select '历史表多余的约束' 差异信息, " & DT_HConDel & " 差异类型, a.Table_Name 表名, a.Constraint_Name 对象名," & vbNewLine & _
            "       f_List2str(Cast(Collect(b.Column_Name Order By b.Position) As t_Strlist)) 列名, Null 数据类型, Null 长度, Null 精度, '×' 自动修复, '手工删除约束' 修复说明" & vbNewLine & _
            "From All_Constraints" & mstrDBLink & " A, All_Cons_Columns" & mstrDBLink & " B" & vbNewLine & _
            "Where a.Owner =  '" & UCase(mstrBakOwnerName) & "' And a.Owner = b.Owner And a.Constraint_Name = b.Constraint_Name And Exists" & vbNewLine & _
            " (Select 1 From Zltools.Zlbaktables T Where t.系统 =  " & mlngSys & " And t.表名 = a.Table_Name) And" & vbNewLine & _
            "      (a.Constraint_Type Not In ('P', 'U') Or a.Constraint_Type In ('P', 'U') And Not Exists" & vbNewLine & _
            "       (Select 1 From User_Constraints C Where c.Constraint_Name = a.Constraint_Name))" & vbNewLine & _
            "Group By a.Table_Name, a.Constraint_Name"

    Call OpenRecordset(rsTmp, strSQL, "历史表多余的约束")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '（4）历史表空间缺少的主键或唯一键约束
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表缺少的约束")
    strSQL = "Select  '历史表缺少的约束' 差异信息, " & DT_HConAdd & " 差异类型, Table_Name 表名, Constraint_Name 对象名," & vbNewLine & _
            "       f_List2str(Cast(Collect(Column_Name Order By a.Position) As t_Strlist)) 列名, Constraint_Type 数据类型, Null 长度, Null 精度, '√' 自动修复," & vbNewLine & _
            "       '创建约束' 修复说明" & vbNewLine & _
            "From (Select a.Table_Name, a.Constraint_Name, a.Column_Name, Nvl(a.Position, 1) Position, b.Constraint_Type" & vbNewLine & _
            "       From User_Cons_Columns A, User_Constraints B, Zltools.Zlbaktables T" & vbNewLine & _
            "       Where a.Constraint_Name = b.Constraint_Name And b.Table_Name = t.表名 And t.系统 = " & mlngSys & " And" & vbNewLine & _
            "             b.Constraint_Type In ('P', 'U') And Not Exists" & vbNewLine & _
            "        (Select 1" & vbNewLine & _
            "              From All_Constraints" & mstrDBLink & " C" & vbNewLine & _
            "              Where c.Owner = '" & UCase(mstrBakOwnerName) & "' And c.Constraint_Type In ('P', 'U') And c.Table_Name = t.表名 And" & vbNewLine & _
            "                    c.Constraint_Name = b.Constraint_Name)) A" & vbNewLine & _
            "Group By a.Table_Name, a.Constraint_Name, Constraint_Type"
            
    Call OpenRecordset(rsTmp, strSQL, "历史表缺少的约束")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '(5)历史表空间约束列变动
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表约束列的差异")
    strSQL = "Select  '约束列差异' 差异信息, " & DT_HConColDIf & " 差异类型, u.Table_Name 表名, u.Constraint_Name 对象名, u.Colstr 列名," & vbNewLine & _
            "       (Select c.Constraint_Type" & vbNewLine & _
            "         From User_Constraints C" & vbNewLine & _
            "         Where c.Constraint_Name = u.Constraint_Name And u.Table_Name = c.Table_Name) 数据类型, Null 长度, Null 精度, '√' 自动修复," & vbNewLine & _
            "       '重建约束' 修复说明" & vbNewLine & _
            "From (Select a.Table_Name, a.Constraint_Name, f_List2str(Cast(Collect(a.Column_Name Order By a.Position) As t_Strlist)) Colstr" & vbNewLine & _
            "       From User_Cons_Columns A, Zltools.Zlbaktables T" & vbNewLine & _
            "       Where a.Table_Name = t.表名 And t.系统 = " & mlngSys & vbNewLine & _
            "       Group By a.Table_Name, a.Constraint_Name) U," & vbNewLine & _
            "     (Select a.Table_Name, a.Constraint_Name, f_List2str(Cast(Collect(a.Column_Name Order By a.Position) As t_Strlist)) Colstr" & vbNewLine & _
            "       From All_Cons_Columns" & mstrDBLink & " A, Zltools.Zlbaktables T" & vbNewLine & _
            "       Where a.Table_Name = t.表名 And t.系统 = " & mlngSys & " And a.Owner =  '" & UCase(mstrBakOwnerName) & "'" & vbNewLine & _
            "       Group By a.Table_Name, a.Constraint_Name) H" & vbNewLine & _
            "Where u.Table_Name = h.Table_Name And u.Constraint_Name = h.Constraint_Name And u.Colstr <> h.Colstr"

    '(5)历史表索引表空间错误
    lngCur = lngCur + 1
    SetPromptText ("（" & lngCur & "/" & lngTotal & ")正在检查历史表约束列的差异")
    strSQL = "Select '历史表索引表空间错误' 差异信息, " & DT_HIndexTablesapce & " 差异类型, a.Table_Name 表名, a.Index_Name 对象名, Null 列名, Null 数据类型, Null 长度, Null 精度, '√' 自动修复," & vbNewLine & _
            "       '移动索引到表空间" & mstrBakIndexDB & "' 修复说明" & vbNewLine & _
            "From (Select d.Table_Name, d.Index_Name," & vbNewLine & _
            "              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
            "       From User_Ind_Columns d, Zltools.Zlbaktables t" & vbNewLine & _
            "       Where d.Table_Name = t.表名" & vbNewLine & _
            "       And t.系统 = " & mlngSys & "" & vbNewLine & _
            "       And Instr(d.Index_Name, '_IX_待转出') = 0" & vbNewLine & _
            "       Group By d.Table_Name, d.Index_Name) a," & vbNewLine & _
            "     (Select e.Table_Name, f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr" & vbNewLine & _
            "       From User_Cons_Columns e, User_Constraints f, Zltools.Zlbaktables t, User_Constraints c," & vbNewLine & _
            "            (Select 表名 From Zltools.Zlbaktables Union All" & vbNewLine & _
            "              Select '病案主页' From Dual Union All" & vbNewLine & _
            "              Select '病人信息' From Dual) g" & vbNewLine & _
            "       Where e.Table_Name = t.表名 And t.系统 = " & mlngSys & " And e.Constraint_Name = f.Constraint_Name" & vbNewLine & _
            "       And f.Constraint_Type = 'R' And c.Constraint_Name = f.r_Constraint_Name And g.表名(+) = c.Table_Name And g.表名 Is Null" & vbNewLine & _
            "       Group By e.Table_Name, e.Constraint_Name) b," & vbNewLine & _
            "     (Select Index_Name From All_Indexes" & mstrDBLink & " k, Zltools.Zlbaktables t" & vbNewLine & _
            "       Where k.Table_Owner = '" & UCase(mstrBakOwnerName) & "' And k.Table_Name = t.表名 And t.系统 = " & mlngSys & vbNewLine & _
            "       And k.Status = 'VALID'  And k.Tablespace_Name Not in( '" & mstrBakIndexDB & "'" & IIf(mstrBakIndexDB = mstrBakDB, ",'" & mstrBakDB & "_IDX')", ")") & ") h" & vbNewLine & _
            "Where a.Table_Name = b.Table_Name(+) And a.Colstr = b.Colstr(+)" & vbNewLine & _
            "And (b.Colstr Is Null Or Instr(a.Index_Name, '_PK') > 0 Or Instr(a.Index_Name, '_UQ') > 0)" & vbNewLine & _
            "And a.Index_Name = h.Index_Name"
            
    'strLobTable对应的调整SQL
    'k.Status = 'VALID'  And  (k.Tablespace_Name Not in( '" & mstrBakIndexDB & "'" & IIf(mstrBakIndexDB = mstrBakDB, ",'" & mstrBakDB & "_IDX')", ")") & " OR K.Table_Name In(" & strLobTable & "))) h" & vbNewLine & _

    Call OpenRecordset(rsTmp, strSQL, "历史表约束列差异")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    
    If Not mblnUpdate Then
        LoadCheckData = vsCheckResult.Rows <> vsCheckResult.FixedRows
        
        With vsCheckResult
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, RC_DifInfo) = "H视图重建"
            .TextMatrix(.Rows - 1, RC_AutoRep) = "√"
            .TextMatrix(.Rows - 1, RC_RepMethod) = "重新创建H视图"
             .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, RC_DifInfo) = "相关过程重新编译"
            .TextMatrix(.Rows - 1, RC_AutoRep) = "√"
            .TextMatrix(.Rows - 1, RC_RepMethod) = "重新编译数据转移相关的存储过程"
        End With
        
        vsCheckResult.Redraw = True
    End If
    Exit Function
errH:
    If 1 = 0 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Function


Private Sub LoadDataByRecord(ByVal rsTmp As ADODB.Recordset, Optional ByVal blnClear As Boolean)
    '-------------------------------------------------------------------------------------------------------------
    '功能:将记录集加载到历史表空间检查表格中
    '参数：rsTmp 检查结果记录集
    '      blnClear 清空表格内容
    '-------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errH:
    With vsCheckResult
    
        If blnClear Then
            .Rows = .FixedRows
        End If
        
        If rsTmp.RecordCount = 0 Then
            Exit Sub
        End If
        
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            
            If Val(rsTmp!差异类型) = -1 Then
                .TextMatrix(lngRow, RC_ObjType) = rsTmp!数据类型 & ""
                .TextMatrix(lngRow, RC_ObjLen) = rsTmp!后备长度 & ""
                .TextMatrix(lngRow, RC_ObjScale) = rsTmp!后备精度 & ""
                .Cell(flexcpData, lngRow, RC_ObjLen) = rsTmp!在线长度 & ""
                .Cell(flexcpData, lngRow, RC_ObjScale) = rsTmp!在线精度 & ""
                .TextMatrix(lngRow, RC_DifInfo) = "列精度差异"
                If Val(rsTmp!后备精度 & "") <= Val(rsTmp!在线精度 & "") And (Val(rsTmp!后备长度 & "") - Val(rsTmp!后备精度 & "")) <= (Val(rsTmp!在线长度 & "") - Val(rsTmp!在线精度 & "")) Then
                    .TextMatrix(lngRow, RC_DifType) = 4
                    .TextMatrix(lngRow, RC_AutoRep) = "√"
                    .TextMatrix(lngRow, RC_RepMethod) = "扩大精度"
                    
                Else
                    .TextMatrix(lngRow, RC_DifType) = 5
                    .TextMatrix(lngRow, RC_AutoRep) = "×"
                    .TextMatrix(lngRow, RC_RepMethod) = "历史表空间该列精度大于在线库,可能会影响抽选返回功能"
                End If
            Else
                .TextMatrix(lngRow, RC_ObjType) = rsTmp!数据类型 & ""
                .TextMatrix(lngRow, RC_ObjLen) = rsTmp!长度 & ""
                .TextMatrix(lngRow, RC_ObjScale) = rsTmp!精度 & ""
                .TextMatrix(lngRow, RC_DifType) = rsTmp!差异类型 & ""
                .TextMatrix(lngRow, RC_DifInfo) = rsTmp!差异信息 & ""
                .TextMatrix(lngRow, RC_AutoRep) = rsTmp!自动修复 & ""
                .TextMatrix(lngRow, RC_RepMethod) = rsTmp!修复说明 & ""
            End If
            .TextMatrix(lngRow, RC_TabName) = rsTmp!表名 & ""
            .TextMatrix(lngRow, RC_ObjName) = rsTmp!对象名 & ""
            .TextMatrix(lngRow, RC_ColName) = rsTmp!列名 & ""
            
            If .TextMatrix(lngRow, RC_DifType) <> "" Then
                Select Case Val(.TextMatrix(lngRow, RC_DifType))
                    Case DT_HLackTab
                        If mstrDBLink = "" Then
                            ReDim Preserve marrOnlineAddSQL(UBound(marrOnlineAddSQL) + 1)
                            marrOnlineAddSQL(UBound(marrOnlineAddSQL)) = " Grant Select On " & .TextMatrix(lngRow, RC_TabName) & " To " & mstrBakOwnerName  '对后备库管理员授予在线库相应表的Select权限
                            If Not ExistsSynonym(.TextMatrix(lngRow, RC_TabName)) Then  '为表创建公共同义词
                                ReDim Preserve marrOnlineAddSQL(UBound(marrOnlineAddSQL) + 1)
                                marrOnlineAddSQL(UBound(marrOnlineAddSQL)) = " Create Public Synonym " & .TextMatrix(lngRow, RC_TabName) & " For " & .TextMatrix(lngRow, RC_TabName)
                            End If
                        End If
                        strSQL = CreateTable(gcnOracle, mstrOwnerName, mstrBakDB, mstrBakOwnerName, .TextMatrix(lngRow, RC_TabName), mstrBakLobDB)
                        If strSQL <> "" Then
                            .TextMatrix(lngRow, RC_RepSQL) = strSQL
                        End If
                        If mstrDBLink = "" Then
                            '对在线库管理员授予后备库相应表的所有权限
                            ReDim Preserve marrBakAddSQL(UBound(marrBakAddSQL) + 1)
                            marrBakAddSQL(UBound(marrBakAddSQL)) = " Grant All On " & .TextMatrix(lngRow, RC_TabName) & " To " & mstrOwnerName & " with Grant option"
                        End If
                    Case DT_HMoreCol '历史表空间多一列
                    
                    Case DT_HLessCol '历史表空间少一列
                        If Val(.TextMatrix(lngRow, RC_ObjLen)) = 0 Then
                            strTmp = .TextMatrix(lngRow, RC_ObjType)
                        Else
                            If Val(.TextMatrix(lngRow, RC_ObjScale)) = 0 Then
                                strTmp = .TextMatrix(lngRow, RC_ObjType) & "(" & .TextMatrix(lngRow, RC_ObjLen) & ")"
                            Else
                                strTmp = .TextMatrix(lngRow, RC_ObjType) & "(" & .TextMatrix(lngRow, RC_ObjLen) & "," & Val(.TextMatrix(lngRow, RC_ObjScale)) & ")"
                            End If
                        End If
                        .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Add " & .TextMatrix(lngRow, RC_ColName) & " " & strTmp
                    Case DT_HDataTypeDif '列类型差异
                        
                    Case DT_HRepLenDif '可修复列精度差异
                        If Val(.Cell(flexcpData, lngRow, RC_ObjLen)) = 0 Then
                            strTmp = .TextMatrix(lngRow, RC_ObjType)
                        Else
                            If Val(.Cell(flexcpData, lngRow, RC_ObjScale)) = 0 Then
                                strTmp = .TextMatrix(lngRow, RC_ObjType) & "(" & .Cell(flexcpData, lngRow, RC_ObjLen) & ")"
                            Else
                                strTmp = .TextMatrix(lngRow, RC_ObjType) & "(" & .Cell(flexcpData, lngRow, RC_ObjLen) & "," & Val(.Cell(flexcpData, lngRow, RC_ObjScale)) & ")"
                            End If
                        End If
                        .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Modify " & .TextMatrix(lngRow, RC_ColName) & " " & strTmp
                    Case DT_HNotRepLenDif '不可修复列精度差异
                    
                    Case DT_HIndUsable, DT_HIndexTablesapce '历史表失效索引
                        .TextMatrix(lngRow, RC_RepSQL) = "Alter Index " & .TextMatrix(lngRow, RC_ObjName) & " Rebuild Tablespace " & mstrBakIndexDB
                    Case DT_HIndDel '历史表多余的索引
                        .TextMatrix(lngRow, RC_RepSQL) = "Drop Index " & .TextMatrix(lngRow, RC_ObjName)
                    Case DT_HIndAdd '历史表缺少的索引
                        .TextMatrix(lngRow, RC_RepSQL) = "Create Index " & .TextMatrix(lngRow, RC_ObjName) & " On " & .TextMatrix(lngRow, RC_TabName) & "(" & .TextMatrix(lngRow, RC_ColName) & ")  Tablespace " & mstrBakIndexDB
                    Case DT_HIndColDif  '索引列差异
                        .TextMatrix(lngRow, RC_RepSQL) = "Drop Index " & .TextMatrix(lngRow, RC_ObjName)
                        ReDim Preserve marrBakAddSQL(UBound(marrBakAddSQL) + 1)
                        marrBakAddSQL(UBound(marrBakAddSQL)) = "Create Index " & .TextMatrix(lngRow, RC_ObjName) & " On " & .TextMatrix(lngRow, RC_TabName) & "(" & .TextMatrix(lngRow, RC_ColName) & ")  Tablespace " & mstrBakIndexDB
                    Case DT_HConDisable '历史表中禁用约束
                        .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Enable Constraint " & .TextMatrix(lngRow, RC_ObjName)
                    Case DT_HConDel '历史表多余约束
                        If .TextMatrix(lngRow, RC_ObjName) Like "*_PK" Or .TextMatrix(lngRow, RC_ObjName) Like "*_UQ_*" Then
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Drop Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Cascade Drop Index"
                        Else
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Drop Constraint " & .TextMatrix(lngRow, RC_ObjName)
                        End If
                    Case DT_URefConDel '子表数据未转出
                        '不修复，用户手工修复
                    Case DT_HConAdd '历史表缺少的约束
                        If .TextMatrix(lngRow, RC_ObjType) = "P" Then
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter table " & .TextMatrix(lngRow, RC_TabName) & " Add Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Primary Key(" & .TextMatrix(lngRow, RC_ColName) & ")  Using Index  Tablespace " & mstrBakIndexDB
                        ElseIf .TextMatrix(lngRow, RC_ObjType) = "U" Then
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter table " & .TextMatrix(lngRow, RC_TabName) & " Add Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Unique(" & .TextMatrix(lngRow, RC_ColName) & ")  Using Index  Tablespace " & mstrBakIndexDB
                        End If
                    Case DT_HConColDIf '约束列差异
                        If .TextMatrix(lngRow, RC_ObjType) = "P" Or .TextMatrix(lngRow, RC_ObjType) = "U" Then
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Drop Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Cascade Drop Index"
                        Else
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Drop Constraint " & .TextMatrix(lngRow, RC_ObjName)
                        End If
                        ReDim Preserve marrBakAddSQL(UBound(marrBakAddSQL) + 1)
                        If .TextMatrix(lngRow, RC_ObjType) = "P" Then
                            marrBakAddSQL(UBound(marrBakAddSQL)) = "Alter table " & .TextMatrix(lngRow, RC_TabName) & " Add Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Primary Key(" & .TextMatrix(lngRow, RC_ColName) & ") Using Index  Tablespace " & mstrBakIndexDB
                        ElseIf .TextMatrix(lngRow, RC_ObjType) = "U" Then
                            marrBakAddSQL(UBound(marrBakAddSQL)) = "Alter table " & .TextMatrix(lngRow, RC_TabName) & " Add Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Unique(" & .TextMatrix(lngRow, RC_ColName) & ") Using Index  Tablespace " & mstrBakIndexDB
                        End If
                    Case DT_HLobTablespace '建议手工处理,若取消屏蔽，注意调整索引表空间检查SQL
    '                    .TextMatrix(lngRow, RC_RepSQL) ="Alter Table " & .TextMatrix(lngRow, RC_TabName)名 & " Move Tablespace " & mstrBakLobDB
                End Select
            End If
            If .TextMatrix(lngRow, RC_AutoRep) = "×" Then .Cell(flexcpBackColor, lngRow, .FixedCols, lngRow, .Cols - 1) = &H8000000F
            rsTmp.MoveNext
        Wend
    End With
    Exit Sub
errH:
    If 1 = 0 Then
        Resume
    End If
End Sub

Private Function ExistsSynonym(ByVal strTableName As String) As Boolean
'功能:查询当前表是否存在公共同义词
'       strTableName 要检查的表
'返回：true-存在公共同义词，false-不存在公共同义词

    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    strSQL = "Select 1 From All_Synonyms A Where a.Table_Owner = User And a.Synonym_Name ='" & strTableName & "' And a.Owner = 'PUBLIC'"
    Call OpenRecordset(rsTmp, strSQL, Me.Caption)
    ExistsSynonym = Not rsTmp.EOF
    
End Function

Private Sub LoadErrInfo(ByRef rsTmp As ADODB.Recordset)
    Dim lngRow As Long, i As Long

    With vsCheckResult
        .Rows = .FixedRows
        If rsTmp.RecordCount <> 0 Then
            rsTmp.MoveFirst
            For i = 0 To .Cols - 1
                 .ColHidden(i) = True
            Next
            .ColHidden(RC_TabName) = False
            .ColHidden(RC_ObjName) = False
            .ColHidden(RC_DifType) = False
            .ColWidth(RC_DifType) = 3000
            .ColHidden(RC_DifInfo) = False
            .ColHidden(RC_RepSQL) = False
            .ColWidth(RC_RepSQL) = 2700
            .TextMatrix(0, RC_DifInfo) = "数据空间"
            .TextMatrix(0, RC_RepSQL) = "错误SQL"
            .TextMatrix(0, RC_DifType) = "错误信息"
            .TextMatrix(0, RC_RepSQL) = "错误SQL"
            While Not rsTmp.EOF
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                .TextMatrix(lngRow, RC_TabName) = rsTmp!表名 & ""
                .Cell(flexcpData, lngRow, RC_TabName) = .TextMatrix(lngRow, RC_TabName)
                .TextMatrix(lngRow, RC_ObjName) = rsTmp!对象名 & ""
                .Cell(flexcpData, lngRow, RC_ObjName) = .TextMatrix(lngRow, RC_ObjName)
                .TextMatrix(lngRow, RC_DifInfo) = IIf(Val(rsTmp!数据库) = 0, "在线库", "历史表空间")
                .Cell(flexcpData, lngRow, RC_DifInfo) = .TextMatrix(lngRow, RC_DifInfo)
                .TextMatrix(lngRow, RC_DifType) = rsTmp!差异信息 & ""
                .Cell(flexcpData, lngRow, RC_DifType) = .TextMatrix(lngRow, RC_DifType)
                .TextMatrix(lngRow, RC_RepSQL) = rsTmp!错误SQL & ""
                .Cell(flexcpData, lngRow, RC_RepSQL) = .TextMatrix(lngRow, RC_RepSQL)
                rsTmp.MoveNext
            Wend
        End If
    End With
End Sub

Private Sub SetPromptText(ByVal strText As String)
    If Not mblnUpdate Then
        stbThis.Panels(2).Text = strText
        stbThis.Panels(2).ToolTipText = strText
    End If
End Sub

Private Sub SetFaceCtlEnable()
    cmdExit.Enabled = mblnAllRepair
    cmdRepair.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnAllRepair Then
        Cancel = True
    Else
        Set mrsErrInfo = Nothing
        Set mcnBakDB = Nothing
    End If
End Sub

Private Sub vsCheckResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsCheckResult.TextMatrix(Row, Col) = vsCheckResult.Cell(flexcpData, Row, Col)
End Sub

Private Sub vsCheckResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsCheckResult.FixedRows And NewCol >= vsCheckResult.FixedCols Then
        If NewRow <> OldRow Then
            vsCheckResult.ForeColorSel = &H0&
        End If
    End If
End Sub

Private Sub AddErrIntoRs(ByVal intDBType As Integer, Optional ByVal strErrInfo As String, Optional ByVal strTabName As String, Optional ByVal strObjName As String, Optional ByVal strSQL As String)
'功能：加载错误信息于记录集
    With mrsErrInfo
        .AddNew
        !数据库 = intDBType
        If InStr(strErrInfo, "ORA-") > 0 Then
            !差异信息 = Mid(strErrInfo, InStr(strErrInfo, "ORA-"))
        Else
            !差异信息 = strErrInfo
        End If
        
        !表名 = strTabName
        !对象名 = strObjName
        !错误SQL = strSQL
        .Update
    End With
End Sub

Private Sub vsCheckResult_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = cmdRepair.Enabled Or Row < vsCheckResult.FixedRows
    If Not Cancel Then Cancel = Col <> RC_RepSQL
End Sub

Private Function GetIniRec() As ADODB.Recordset
'功能：获取初始SQL记录集，升级程序调用时，需返回历史库修正的SQL记录集，此处做SQL记录集初始化
    Dim rsReturn As New ADODB.Recordset
    
    With rsReturn
        .Fields.Append "BAKDBName", adVarChar, 100 '历史库名称
        .Fields.Append "BAKUser", adVarChar, 100
        .Fields.Append "SQL", adVarChar, 500 '数据库修复SQL
        .Fields.Append "ExecOrder", adInteger '用于确定SQL执行的前后顺序，有些修改需要一些权限，这些权限SQL需要提前执行，修复完成后有事需要一些后续处理，这些处理要稍后执行。
        .Fields.Append "FixType", adInteger '修复类型，用来区分SQL执行顺序
        .Fields.Append "ExecDB", adInteger '执行SQL的数据库，0-历史库，1-在线库
        .Fields.Append "ExecIndex", adInteger 'SQL加载顺序
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    Set GetIniRec = rsReturn
End Function

Private Function GetFixSQL(ByVal rsInput As ADODB.Recordset, ByRef rsSQL As ADODB.Recordset)
'功能：根据查询出的差异信息获取修正SQL
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String
    
    With rsInput
        While Not .EOF
            Select Case Val(!差异类型)
                Case DT_HLackTab
                    '提前执行SQL,ExecOrder=-1
                    If mstrDBLink = "" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, " Grant Select On " & !表名 & " To " & mstrBakOwnerName, -1, DT_HLackTab, 1, mlngIndex)
                        mlngIndex = mlngIndex + 1
                    End If
                    '正常执行SQL,ExecOrder=0
                    strSQL = CreateTable(gcnOracle, mstrOwnerName, mstrBakDB, mstrBakOwnerName, !表名, mstrBakLobDB)
                    If strSQL <> "" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, strSQL, 0, DT_HLackTab, 0, mlngIndex)
                        mlngIndex = mlngIndex + 1
                    End If
                    If mstrDBLink = "" Then
                        '延后执行SQL,ExecOrder=1
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, " Grant All On " & !表名 & " To " & mstrOwnerName & " with Grant option", 1, DT_HLackTab, 0, mlngIndex)
                        mlngIndex = mlngIndex + 1
                    End If
                Case DT_HMoreCol
                Case DT_HLessCol
                    If Val(!长度 & "") = 0 Then
                        strTmp = !数据类型 & ""
                    Else
                        If Val(!精度 & "") = 0 Then
                            strTmp = !数据类型 & "(" & !长度 & ")"
                        Else
                            strTmp = !数据类型 & "(" & !长度 & "," & !精度 & ")"
                        End If
                    End If
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !表名 & " Add " & !列名 & " " & strTmp, 0, DT_HLessCol, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HDataTypeDif
                
                Case -1
                    'Case DT_HRepLenDif
                    If Val(!后备精度 & "") <= Val(!在线精度 & "") And (Val(!后备长度 & "") - Val(!后备精度 & "")) <= (Val(!在线长度 & "") - Val(!在线精度 & "")) Then
                        If Val(!在线长度 & "") = 0 Then
                            strTmp = !数据类型 & ""
                        Else
                            If Val(!在线精度 & "") = 0 Then
                                strTmp = !数据类型 & "(" & !在线长度 & ")"
                            Else
                                strTmp = !数据类型 & "(" & !在线长度 & "," & !在线精度 & ")"
                            End If
                        End If
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !表名 & " Modify " & !列名 & " " & strTmp, 0, DT_HRepLenDif, 0, mlngIndex)
                        mlngIndex = mlngIndex + 1
                    Else
                    'Case DT_HNotRepLenDif
                    
                    End If
                Case DT_HIndUsable, DT_HIndexTablesapce
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Index " & !对象名 & " Rebuild TableSpace " & mstrBakIndexDB, 0, DT_HIndUsable, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HIndDel
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Drop Index " & !对象名, 0, DT_HIndDel, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HIndAdd
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Create Index " & !对象名 & " On " & !表名 & "(" & !列名 & ") TableSpace " & mstrBakIndexDB, 0, DT_HIndAdd, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HIndColDif
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Drop Index " & !对象名, 0, DT_HIndColDif, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Create Index " & !对象名 & " On " & !表名 & "(" & !列名 & ")  TableSpace " & mstrBakIndexDB, 0, DT_HIndColDif, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HConDisable
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !表名 & " Enable Constraint " & !对象名, 0, DT_HConDisable, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HConDel
                    If !对象名 Like "*_PK" Or !对象名 Like "*_UQ_*" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !表名 & " Drop Constraint " & !对象名 & " Cascade Drop Index", 0, DT_HConDel, 0, mlngIndex)
                    Else
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !表名 & " Drop Constraint " & !对象名, 0, DT_HConDel, 0, mlngIndex)
                    End If
                    mlngIndex = mlngIndex + 1
                Case DT_URefConDel
                Case DT_HConAdd '可能缺少外键约束
                    '检查索引是否存在，存在，则删除
                    strSQL = "Select /*+rule*/" & vbNewLine & _
                                " 1" & vbNewLine & _
                                "From User_Indexes A" & vbNewLine & _
                                "Where A.Index_Name = '" & !对象名 & "'"
                    Call OpenRecordset(rsTmp, strSQL, "数据转移相关过程有效检查", , , mcnBakDB)
                    If Not rsTmp.EOF Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, " Drop Index  " & !对象名, -1, DT_HConAdd, 0, mlngIndex)
                    End If
                    If !数据类型 & "" = "P" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter table " & !表名 & " Add Constraint " & !对象名 & " Primary Key(" & !列名 & ") Using Index  TableSpace " & mstrBakIndexDB, 0, DT_HConAdd, 0, mlngIndex)
                    ElseIf !数据类型 & "" = "U" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter table " & !表名 & " Add Constraint " & !对象名 & " Unique(" & !列名 & ")  Using Index  TableSpace " & mstrBakIndexDB, 0, DT_HConAdd, 0, mlngIndex)
                    End If
                    mlngIndex = mlngIndex + 1
                Case DT_HConColDIf
                    If !数据类型 & "" = "P" Or !数据类型 & "" = "U" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !表名 & " Drop Constraint " & !对象名 & " Cascade Drop Index", 0, DT_HConAdd, 0, mlngIndex)
                    Else
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !表名 & " Drop Constraint " & !对象名, 0, DT_HConAdd, 0, mlngIndex)
                    End If
                    mlngIndex = mlngIndex + 1
                    If !数据类型 & "" = "P" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter table " & !表名 & " Add Constraint " & !对象名 & " Primary Key(" & !列名 & ")  Using Index  TableSpace " & mstrBakIndexDB, 0, DT_HConAdd, 0, mlngIndex)
                    ElseIf !数据类型 & "" = "U" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter table " & !表名 & " Add Constraint " & !对象名 & " Unique(" & !列名 & ")  Using Index  TableSpace " & mstrBakIndexDB, 0, DT_HConAdd, 0, mlngIndex)
                    End If
                    mlngIndex = mlngIndex + 1
                Case DT_HLobTablespace '建议手工处理,若取消屏蔽，注意调整索引表空间检查SQL
'                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !表名 & " Move Tablespace " & mstrBakLobDB, 0, DT_HConAdd, 0, mlngIndex)
            End Select
            .MoveNext
        Wend
    End With
End Function

Private Sub ADDSQLToRec(ByRef rsSQL As ADODB.Recordset, ByVal strBakDB As String, strBakUser As String, ByVal strSQL As String, ByVal intExecOrder As Integer, ByVal intFixType As Integer, ByVal intExecDB As Integer, ByVal lngExecIndex As Long)
    With rsSQL
        .AddNew
        !BAKDBName = strBakDB
        !BAKUser = strBakUser
        !SQL = strSQL
        !ExecOrder = intExecOrder
        !FixType = intFixType
        !ExecDB = intExecDB
        !ExecIndex = lngExecIndex
        .Update
    End With
End Sub

