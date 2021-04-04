VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAuditItemEdit 
   BorderStyle     =   0  'None
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraAuditItem 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5145
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7965
      Begin VB.Frame fraSource 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   2445
         TabIndex        =   29
         Top             =   60
         Visible         =   0   'False
         Width           =   2970
         Begin VB.OptionButton optSource 
            Caption         =   "标准版"
            Height          =   195
            Index           =   0
            Left            =   1050
            TabIndex        =   31
            Top             =   188
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton optSource 
            Caption         =   "EMR库"
            Height          =   180
            Index           =   1
            Left            =   2070
            TabIndex        =   30
            Top             =   195
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "数据源(&S)"
            Height          =   165
            Left            =   120
            TabIndex        =   32
            Top             =   203
            Width           =   825
         End
      End
      Begin VB.ComboBox CboPalValue 
         Height          =   300
         Left            =   5265
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   555
         Width           =   960
      End
      Begin VB.TextBox txtNumValue 
         DataField       =   "简码"
         Height          =   300
         Left            =   6990
         MaxLength       =   4
         TabIndex        =   25
         ToolTipText     =   "最大扣分值"
         Top             =   570
         Width           =   975
      End
      Begin VB.TextBox txtFileID 
         DataField       =   "文件ID"
         Height          =   300
         Left            =   4710
         TabIndex        =   24
         Top             =   5310
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtCode 
         DataField       =   "编码"
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   22
         Top             =   960
         Width           =   1710
      End
      Begin VB.TextBox txtMnemonicCode 
         DataField       =   "简码"
         Height          =   300
         Left            =   3585
         MaxLength       =   100
         TabIndex        =   21
         Top             =   555
         Width           =   885
      End
      Begin VB.CommandButton cmdSelectFile 
         Height          =   300
         Left            =   6210
         Picture         =   "frmAuditItemEdit.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1965
         Width           =   300
      End
      Begin VB.ComboBox cboLink 
         DataField       =   "适用对象"
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1950
         Width           =   1740
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   300
         Left            =   2865
         Picture         =   "frmAuditItemEdit.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3345
         Width           =   300
      End
      Begin VB.ComboBox cboUsed 
         DataField       =   "适用对象"
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1455
         Width           =   1740
      End
      Begin VB.TextBox txtAudit_NotCheck 
         DataField       =   "审查依据"
         Height          =   780
         Left            =   1050
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         ToolTipText     =   "审查依据"
         Top             =   3345
         Width           =   1740
      End
      Begin VB.TextBox txtName 
         DataField       =   "名称"
         Height          =   300
         Left            =   1065
         MaxLength       =   100
         TabIndex        =   4
         Top             =   570
         Width           =   1710
      End
      Begin VB.TextBox txtDescription 
         DataField       =   "说明"
         Height          =   855
         Left            =   1050
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         ToolTipText     =   "说明"
         Top             =   2385
         Width           =   6915
      End
      Begin VB.TextBox txtTypeID 
         DataField       =   "分类"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   210
         Width           =   6450
      End
      Begin VB.CommandButton cmdSelect 
         Height          =   300
         Left            =   7560
         Picture         =   "frmAuditItemEdit.frx":6930
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   300
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfFiles 
         Height          =   1320
         Left            =   3840
         TabIndex        =   19
         Top             =   915
         Width           =   4200
         _cx             =   7408
         _cy             =   2328
         Appearance      =   2
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAuditItemEdit.frx":D182
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         Editable        =   0
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分制(&M)"
         Height          =   180
         Index           =   8
         Left            =   4560
         TabIndex        =   27
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分值(&N)"
         Height          =   180
         Index           =   4
         Left            =   6285
         TabIndex        =   26
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   3
         Left            =   2850
         TabIndex        =   23
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lab病人ID 
         AutoSize        =   -1  'True
         Caption         =   "[病人ID]、[主页ID]为系统参数，分别代表系统中的病人ID和主页ID。"
         Height          =   180
         Left            =   1515
         TabIndex        =   18
         Top             =   4485
         Width           =   5550
      End
      Begin VB.Label labNote 
         AutoSize        =   -1  'True
         Caption         =   "注："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1065
         TabIndex        =   17
         Top             =   4485
         Width           =   390
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "适用环节(&L)"
         Height          =   180
         Index           =   7
         Left            =   30
         TabIndex        =   8
         Top             =   2010
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "适用对象(&D)"
         Height          =   180
         Index           =   5
         Left            =   30
         TabIndex        =   6
         Top             =   1515
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审查依据(&X)"
         Height          =   180
         Index           =   1
         Left            =   30
         TabIndex        =   12
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label lstFiles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分类(&T)"
         Height          =   180
         Index           =   4
         Left            =   390
         TabIndex        =   0
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明(&Z)"
         Height          =   180
         Index           =   9
         Left            =   390
         TabIndex        =   10
         Top             =   2415
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   3
         Top             =   645
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编码(&C)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   5
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件(&J)"
         Height          =   180
         Index           =   6
         Left            =   3090
         TabIndex        =   16
         Top             =   1500
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmAuditItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mEditMode               As 编辑模式

Private mblnDataChange          As Boolean  '确定 or 取消

Private zlCheck                 As New clsCheck
Private mlngItemID              As Long  'ID
Private mlngItemTypeID          As Long  '分类ID
Private mstrItemCode            As String   '编码
Private mstrItemName            As String   '名称
Private mstrItemMnemonicCode    As String   '简码
Private mstrItemDescription     As String   '说明
Private mstrItemAudit           As String   '审查依据
Private mintItemUsed            As Integer  '适用对象
Private mintItemFileID          As String   '适用对象文件
Private mstrItemLink            As Integer  '适用环节
Private mintUsedListIndex       As Integer  '记录适用对象的序号
Private mstrNumValue            As String   '分值
Private mstrPalValue            As String   '分制
Private mintSource              As Integer  '数据源
Private mblnProgUsed            As Boolean  '方案是否被使用

Private Const conFieldFiles = "Select /*+ rule */ a.id as 文件ID,a.编号 as 文件编码,a.名称 as 文件名称,a.说明 as 文件说明" & vbCrLf & _
                         "from 病历文件列表 A, Table (Cast(f_Str2List([1])  As zlTools.t_StrList)) B " & vbCrLf & _
                         "where /*+ rule */a.id = b.COLUMN_VALUE And a.种类 = [2]"
Private Const conEmrField = "Select /*+ Rule*/ Rawtohex(b.Id) As 文件id, b.Code As 文件编码, b.Title As 文件名称, b.Note As 文件说明" & vbNewLine & _
                        "From (Select Hextoraw(Column_Value) As ID From Table(Zlcommunal.f_Str2list(:p0, ','))) A, Antetype_List B" & vbNewLine & _
                        "Where Hextoraw(a.Id) = b.Id And b.Kind = :p1" & vbNewLine & _
                        "Order By 文件编码"

Public Property Get blnProgUsed() As Boolean
    blnProgUsed = mblnProgUsed
End Property

Public Property Let blnProgUsed(ByVal vNewValue As Boolean)
    mblnProgUsed = vNewValue
End Property

Public Property Get DataChange() As Boolean
    DataChange = mblnDataChange And mEditMode <> 浏览
End Property

Public Property Let DataChange(ByVal vNewValue As Boolean)
    mblnDataChange = vNewValue
End Property

Public Property Get EditMode() As 编辑模式
    EditMode = mEditMode
End Property

Public Property Let EditMode(ByVal vNewValue As 编辑模式)
    mEditMode = vNewValue
End Property

Public Property Get lngItemID() As Long
    lngItemID = mlngItemID
End Property

Public Property Let lngItemID(ByVal vNewValue As Long)
    mlngItemID = vNewValue
End Property

Public Property Get lngItemTypeID() As Long
    lngItemTypeID = mlngItemTypeID
End Property

Public Property Let lngItemTypeID(ByVal vNewValue As Long)
    mlngItemTypeID = vNewValue
End Property

Public Property Get strItemCode() As String
    strItemCode = mstrItemCode
End Property

Public Property Let strItemCode(ByVal vNewValue As String)
    mstrItemCode = vNewValue
End Property

Public Property Get strItemName() As String
    strItemName = mstrItemName
End Property

Public Property Let strItemName(ByVal vNewValue As String)
    mstrItemName = vNewValue
End Property

Public Property Get strItemMnemonicCode() As String
    strItemMnemonicCode = mstrItemMnemonicCode
End Property

Public Property Let strItemMnemonicCode(ByVal vNewValue As String)
    mstrItemMnemonicCode = vNewValue
End Property

Public Property Get strItemDescription() As String
    strItemDescription = mstrItemDescription
End Property

Public Property Let strItemDescription(ByVal vNewValue As String)
    mstrItemDescription = vNewValue
End Property

Public Property Get strItemAudit() As String
    strItemAudit = mstrItemAudit
End Property

Public Property Let strItemAudit(ByVal vNewValue As String)
    mstrItemAudit = vNewValue
End Property

Public Property Get intItemUsed() As Integer
    intItemUsed = mintItemUsed
End Property

Public Property Let intItemUsed(ByVal vNewValue As Integer)
    mintItemUsed = vNewValue
End Property

Public Property Get intItemFileID() As String
    intItemFileID = mintItemFileID
End Property

Public Property Let intItemFileID(ByVal vNewValue As String)
    mintItemFileID = vNewValue
End Property

Public Property Get strItemLink() As String
    strItemLink = mstrItemLink
End Property

Public Property Let strItemLink(ByVal vNewValue As String)
    mstrItemLink = vNewValue
End Property

Public Property Get strNumValue() As String
    strNumValue = mstrNumValue
End Property

Public Property Let strNumValue(ByVal vNewValue As String)
    mstrNumValue = vNewValue
End Property

Public Property Get strPalValue() As String
    strPalValue = mstrPalValue
End Property

Public Property Let strPalValue(ByVal vNewValue As String)
    mstrPalValue = vNewValue
End Property
Public Property Get intSource() As Integer
    intSource = mintSource
End Property

Public Property Let intSource(ByVal vNewValue As Integer)
    mintSource = vNewValue
End Property
Private Sub cboUsed_Click()
    Call cboUsed_LostFocus
End Sub

'==============================================================================
'=功能： 检测
'==============================================================================
Private Sub cmdCheck_Click()
    On Error GoTo ErrH
    If txtAudit_NotCheck.Text <> "" Then Call CheckAuditSql_IN(Trim(txtAudit_NotCheck.Text), True, IIf(optSource(0).Value, 0, 1))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 选择分类
'==============================================================================
Private Sub cmdSelect_Click()
    Dim intTypeID   As Integer
    Dim intLenght   As Integer
    Dim rsTemp      As ADODB.Recordset
    Dim rsTempType      As ADODB.Recordset
    On Error GoTo ErrH
    
    With frmAuditItemTypeSelect
        .lngLeft = frmAuditItem.Left + frmAuditItem.tvwAuditType.Width + txtTypeID.Left + 140
        .lngTop = frmAuditItem.Top + frmAuditItem.vsfAuditItem.Top + frmAuditItem.vsfAuditItem.Height + txtTypeID.Top - 3285
        .intTypeID = IIf(Val(txtTypeID.Tag) = 0, mlngItemTypeID, Val(txtTypeID.Tag))
        .Show vbModal
        If .blnCancel Then Set frmAuditItemTypeSelect = Nothing: Exit Sub
        intTypeID = .intTypeID
    End With
    gstrSQL = "select /*+ rule */id,上级ID,编码,名称 from 病案审查分类 a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(intTypeID))
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        txtTypeID.Tag = CStr(intTypeID)
        txtTypeID.Text = "[" + rsTemp!编码 + "]" & rsTemp!名称
    Else
        txtTypeID.Tag = "-1"
        txtTypeID.Text = "[全部]分类"
    End If
    
    '读取编码
    gstrSQL = "select /*+ rule */id,上级ID,编码,名称 from 病案审查分类 a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtTypeID.Tag)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        gstrSQL = "select nvl(Max(编码),0) from 病案审查目录 where 分类id=[1] and 编码 like [2] || '%'"
        Set rsTempType = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtTypeID.Tag, rsTemp!编码)
        txtCode.Text = IncStr(rsTempType.Fields(0))
    Else
        txtCode.Text = "0001"
    End If
    If txtCode.Text = "1" Then
        txtCode.Text = rsTemp!编码 & Format(txtCode.Text, "0000")
    End If
    txtCode.Text = InsertNewCode(txtCode.Text)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
'==============================================================================
'=功能： 选择文件
'==============================================================================
Private Sub cmdSelectFile_Click()
    Call SelectFile(True)
End Sub

Private Sub SelectFile(Optional blnClick As Boolean = False)
 Dim rsTemp      As ADODB.Recordset, strType As String, strReturn As String
    
    On Error GoTo ErrH
    If cboUsed.Text = "" Or cboUsed.ListIndex = -1 Then
        zlCheck.Msg_OK "请先选择适用对象！"
        Exit Sub
    End If
    strType = AuditFileTran(zlCheck.Cmb_ID(cboUsed), IIf(optSource(0).Value, 0, 1))
    With frmAuditItemEditFile
        .intItemFileID = txtFileID.Text
        .strType = strType
        .intSource = IIf(optSource(0).Value, 0, 1)
        .Show vbModal
        If .intItemFileID = txtFileID.Text Then
            Exit Sub
        End If
         txtFileID.Text = .intItemFileID
    End With
    Set frmAuditItemEditFile = Nothing
    '刷新本地数据
    If optSource(0).Value Then
        gstrSQL = conFieldFiles
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtFileID.Text, strType)
    Else
        gstrSQL = conEmrField
        strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, txtFileID.Text & "^" & DbType.T_String & "^p0|" & strType & "^" & DbType.T_String & "^p1", rsTemp)
        If strReturn <> "" Then
            zlCheck.Msg_OK strReturn
            Exit Sub
        End If
    End If
    Set vsfFiles.DataSource = rsTemp
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub optSource_Click(Index As Integer)
Dim rsUsed As New ADODB.Recordset
    '清除已选文件
    gstrSQL = "Select /*+ rule */" & vbNewLine & _
            " a.Id As 文件id, a.编号 As 文件编码, a.名称 As 文件名称, a.说明 As 文件说明" & vbNewLine & _
            "From 病历文件列表 A" & vbNewLine & _
            "Where a.Id = [1]"
    Set rsUsed = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(0))
    Set vsfFiles.DataSource = rsUsed
    
    If Index = 0 Then
        gstrSQL = "select 1 as ID ,'住院医嘱' as Name from dual union all" & vbCrLf & _
                    "select 2 as ID ,'住院病历' as Name from dual union all" & vbCrLf & _
                    "select 3 as ID ,'护理病历' as Name from dual union all" & vbCrLf & _
                    "select 4 as ID ,'护理记录' as Name from dual union all" & vbCrLf & _
                    "select 5 as ID ,'首页记录' as Name from dual union all" & vbCrLf & _
                    "select 6 as ID ,'医嘱报告' as Name from dual union all" & vbCrLf & _
                    "select 7 as ID ,'疾病证明' as Name from dual union all" & vbCrLf & _
                    "select 8 as ID ,'知情文件' as Name from dual union all" & vbCrLf & _
                    "select 9 as ID ,'临床路径' as Name from dual"
        Set rsUsed = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        zlCheck.Cmb_List cboUsed, rsUsed
        lab病人ID.Caption = "[病人ID]、[主页ID]为系统参数，分别代表系统中的病人ID和主页ID。"
        txtAudit_NotCheck.Tag = 0 '用于放大编辑窗口
    Else
        gstrSQL = "select 2 as ID ,'住院病历' as Name from dual union all" & vbCrLf & _
                    "select 3 as ID ,'护理病历' as Name from dual union all" & vbCrLf & _
                    "select 7 as ID ,'疾病证明' as Name from dual union all" & vbCrLf & _
                    "select 8 as ID ,'知情文件' as Name from dual"
        Set rsUsed = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        zlCheck.Cmb_List cboUsed, rsUsed
        lab病人ID.Caption = "[MID]、[ALIDIN]为系统参数，分别代表EMR中的病人ID和入院ID"
        lab病人ID.ToolTipText = "使用EMR中的ID,除[MID]、[ALIDIN]外都需要使用HextoRaw转换，以便用到索引。"
        txtAudit_NotCheck.Tag = 1 '用于放大编辑窗口
    End If
End Sub
Private Sub txtFileID_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = 13 Then
        Call SelectFile
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 页面初始化
'==============================================================================
Private Sub Form_Load()
    Dim rsUsed      As New ADODB.Recordset
    On Error GoTo ErrH
    mblnDataChange = False
    gstrSQL = "select 1 as ID ,'住院医嘱' as Name from dual union all" & vbCrLf & _
                "select 2 as ID ,'住院病历' as Name from dual union all" & vbCrLf & _
                "select 3 as ID ,'护理病历' as Name from dual union all" & vbCrLf & _
                "select 4 as ID ,'护理记录' as Name from dual union all" & vbCrLf & _
                "select 5 as ID ,'首页记录' as Name from dual union all" & vbCrLf & _
                "select 6 as ID ,'医嘱报告' as Name from dual union all" & vbCrLf & _
                "select 7 as ID ,'疾病证明' as Name from dual union all" & vbCrLf & _
                "select 8 as ID ,'知情文件' as Name from dual union all" & vbCrLf & _
                "select 9 as ID ,'临床路径' as Name from dual"
    Set rsUsed = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    zlCheck.Cmb_List cboUsed, rsUsed
    cboUsed.ListIndex = 0
    cboLink.Clear
    cboLink.AddItem "0" & strSplitCmb & "全部"
    cboLink.AddItem "1" & strSplitCmb & "审查"
    cboLink.AddItem "2" & strSplitCmb & "抽查"
    
    CboPalValue.Clear
    CboPalValue.AddItem "扣分制": CboPalValue.ItemData(CboPalValue.NewIndex) = 0
    CboPalValue.AddItem "否决制": CboPalValue.ItemData(CboPalValue.NewIndex) = 1
    CboPalValue.ListIndex = 0
    
    '字段宽度
    Set rsUsed = zlCheck.GetRsFieldWidth("病案审查目录")
    rsUsed.Filter = "列名='" & txtName.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtName.MaxLength = "" & rsUsed.Fields("长度")
    rsUsed.Filter = "列名='" & txtCode.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtCode.MaxLength = "" & rsUsed.Fields("长度")
    rsUsed.Filter = "列名='" & txtMnemonicCode.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtMnemonicCode.MaxLength = "" & rsUsed.Fields("长度")
    rsUsed.Filter = "列名='" & txtDescription.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtDescription.MaxLength = "" & rsUsed.Fields("长度")
    rsUsed.Filter = "列名='" & txtAudit_NotCheck.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtAudit_NotCheck.MaxLength = "" & rsUsed.Fields("长度")
    If gobjEmr Is Nothing Then
        fraSource.Visible = False
        optSource(0).Value = True
        optSource(1).Value = False
    Else
        fraSource.Visible = True
        optSource(0).Value = False
        optSource(1).Value = True
    End If
    
    zlCheck.Sys_System Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 窗口位置调整
'==============================================================================
Private Sub Form_Resize()
    On Error Resume Next
    fraAuditItem.Move 0, -75, Me.ScaleWidth, Me.ScaleHeight + 75
    txtTypeID.Move txtTypeID.Left, txtTypeID.Top, fraAuditItem.Width - txtTypeID.Left - 500 - IIf(gobjEmr Is Nothing, 0, fraSource.Width), txtTypeID.Height
    cmdSelect.Move txtTypeID.Left + txtTypeID.Width + 60, cmdSelect.Top, cmdSelect.Width, cmdSelect.Height
    fraSource.Move cmdSelect.Left + cmdSelect.Width
    
    txtMnemonicCode.Move txtTypeID.Left + txtTypeID.Width + IIf(gobjEmr Is Nothing, 0, fraSource.Width + 100) - txtMnemonicCode.Width - txtNumValue.Width - lbl(4).Width - lbl(8).Width - CboPalValue.Width - 200 '- txtNumValue.Width - 1000
    lbl(3).Move txtMnemonicCode.Left - lbl(3).Width - 50
    txtName.Move txtName.Left, txtName.Top, lbl(3).Left - txtTypeID.Left - 75, txtTypeID.Height
    
    
    lbl(8).Move txtMnemonicCode.Left + txtMnemonicCode.Width + 50
    CboPalValue.Move lbl(8).Left + lbl(8).Width + 50
    
    lbl(4).Move CboPalValue.Left + CboPalValue.Width + 50


    txtDescription.Move txtDescription.Left, txtDescription.Top, txtTypeID.Width + IIf(gobjEmr Is Nothing, 0, fraSource.Width + 100) + 50, txtDescription.Height
    txtAudit_NotCheck.Move txtAudit_NotCheck.Left, txtAudit_NotCheck.Top, txtTypeID.Width + IIf(gobjEmr Is Nothing, 0, fraSource.Width + 100) + 50, fraAuditItem.Height - txtAudit_NotCheck.Top - 300
    labNote.Top = txtAudit_NotCheck.Top + txtAudit_NotCheck.Height + 50
    lab病人ID.Top = labNote.Top
    cmdCheck.Move txtAudit_NotCheck.Left + txtAudit_NotCheck.Width + 50
    
    vsfFiles.Move vsfFiles.Left, txtCode.Top, txtTypeID.Width + IIf(gobjEmr Is Nothing, 0, fraSource.Width + 100) - vsfFiles.Left + txtAudit_NotCheck.Left + 30
    cmdSelectFile.Move vsfFiles.Left + vsfFiles.Width + 60, vsfFiles.Top, cmdSelectFile.Width, cmdSelectFile.Height
    
     txtNumValue.Move lbl(4).Left + lbl(4).Width + 50, txtMnemonicCode.Top, 800 ', txtTypeID.Left + txtTypeID.Width - (lbl(4).Left + lbl(4).Width + 50), txtTypeID.Height
End Sub
 
'==============================================================================
'=功能： 锁定文本框
'==============================================================================
Public Sub WinLock()
    Dim ctrAll           As Control
    
    On Error GoTo ErrH
    For Each ctrAll In Me.Controls
        If TypeName(ctrAll) = "TextBox" Then
            ctrAll.Locked = True
            ctrAll.BackColor = &H80000000
        ElseIf TypeName(ctrAll) = "CommandButton" Then
            ctrAll.Enabled = False
        ElseIf TypeName(ctrAll) = "ComboBox" Then
            ctrAll.Locked = True
            ctrAll.BackColor = &H80000000
        End If
    Next
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 初始菜单工具栏
'==============================================================================
Public Sub unWinLock(Optional blnCopy As Boolean)
    Dim ctrAll           As Control
    
    On Error GoTo ErrH
    
    For Each ctrAll In Me.Controls
        If TypeName(ctrAll) = "TextBox" Then
            If Not blnCopy Then ctrAll.Text = ""
            ctrAll.Locked = False
            ctrAll.BackColor = vbWhite
        ElseIf TypeName(ctrAll) = "CommandButton" Then
            ctrAll.Enabled = True
        ElseIf TypeName(ctrAll) = "ComboBox" Then
            ctrAll.Locked = False
            ctrAll.BackColor = vbWhite
        End If
    Next
    txtTypeID.Locked = True
    txtTypeID.BackColor = &H80000000
    
    txtFileID.Locked = True
    txtFileID.BackColor = &H80000000
    
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub AuditItemItemInsert()
    On Error GoTo ErrH
    mlngItemID = zlDatabase.GetNextId("病案审查目录")
    gstrSQL = "Zl_病案审查目录_Insert (" + CStr(mlngItemID) + "," + IIf(mlngItemTypeID = 0, "NULL", CStr(mlngItemTypeID)) + "," + "'" + mstrItemCode + "'" + "," + "'" + mstrItemName + "','" & mstrItemMnemonicCode & "','" & mstrItemDescription & "','" & mstrItemAudit & "'," & mintItemUsed & ",'" & mintItemFileID & "','" & mstrItemLink & "'," & Val(mstrPalValue) & " ," & Val(mstrNumValue) & "," & mintSource & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub AuditItemItemUpdate()
    On Error GoTo ErrH
    
    gstrSQL = "Zl_病案审查目录_Update (" + CStr(mlngItemID) + "," + IIf(mlngItemTypeID = 0, "NULL", CStr(mlngItemTypeID)) + "," + "'" + mstrItemCode + "'" + "," + "'" + mstrItemName + "','" & mstrItemMnemonicCode & "','" & mstrItemDescription & "','" & mstrItemAudit & "'," & mintItemUsed & ",'" & mintItemFileID & "','" & mstrItemLink & "'," & Val(mstrPalValue) & ", " & Val(mstrNumValue) & "," & mintSource & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub AuditItemItemDelete()
    On Error GoTo ErrH
    Dim rsTmp As New ADODB.Recordset
    
    If blnProgUsed Then '如果该方案被使用,需要检查项目是否已经被用,如果被使用不能进行删除.
        Set rsTmp = gclsPackage.GetItemUse(mlngItemID)
        If rsTmp.RecordCount = 1 Then
            If rsTmp!条数 > 0 Then
                Call MsgBox("该方项目经被使用过,暂时不能被删除!", vbInformation, gstrSysName)
                Exit Sub
            End If
        End If
    End If
    
    gstrSQL = "Zl_病案审查目录_Delete (" & mlngItemID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function InsertNewCode(strInCode) As String
    Dim rsTemp          As ADODB.Recordset
    On Error GoTo ErrH
    
    gstrSQL = "select 1 from 病案审查目录 where 编码 = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strInCode)
    If zlCheck.Connection_ChkRsState(rsTemp) Then InsertNewCode = strInCode: Exit Function
    strInCode = IncStr(strInCode)
    InsertNewCode = InsertNewCode(strInCode)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 复制添加项目数据 ItemCopy
'==============================================================================
Public Sub ItemCopy()
    Dim strEditMode     As String
    Dim rsTemp          As ADODB.Recordset
    Dim rsTempType      As ADODB.Recordset
    On Error GoTo ErrH
    mEditMode = 复制新增
    unWinLock True
    '读取编码
    gstrSQL = "select /*+ rule */id,上级ID,编码,名称 from 病案审查分类 a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtTypeID.Tag)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        gstrSQL = "select nvl(Max(编码),0) from 病案审查目录 where 分类id=[1] and 编码 like [2] || '%'"
        Set rsTempType = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtTypeID.Tag, rsTemp!编码)
        txtCode.Text = IncStr(rsTempType.Fields(0))
    Else
        txtCode.Text = "0001"
    End If
    If txtCode.Text = "1" Then
        txtCode.Text = rsTemp!编码 & Format(txtCode.Text, "0000")
    End If
    txtCode.Text = InsertNewCode(txtCode.Text)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 添加项目数据 ItemInsert
'==============================================================================
Public Sub ItemInsert()
    Dim strEditMode     As String
    Dim strID           As String
    Dim rsTemp          As ADODB.Recordset
    On Error GoTo ErrH
    strEditMode = ""
    
    mEditMode = 新增
    unWinLock
    
    gstrSQL = "select /*+ rule */id,上级ID,编码,名称 from 病案审查分类 a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngItemTypeID)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        txtTypeID.Tag = strID
        txtTypeID.Text = "[" + rsTemp!编码 + "]" & rsTemp!名称
    Else
        txtTypeID.Tag = "-1"
        txtTypeID.Text = "[全部]分类"
    End If
    txtName.SetFocus
    '读取编码
    
    CboPalValue.ListIndex = 0
    
    gstrSQL = "select nvl(Max(编码),0) from 病案审查目录 where 分类id=[1] and 编码 like [2] || '%'"
    txtCode.Text = IncStr(zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngItemTypeID, rsTemp!编码).Fields(0))
    If txtCode.Text = "1" Then
        txtCode.Text = rsTemp!编码 & Format(txtCode.Text, "0000")
    End If
    txtCode.Text = InsertNewCode(txtCode.Text)
    If gobjEmr Is Nothing Then
        optSource(0).Value = True
        optSource(1).Value = False
    Else
        optSource(0).Value = False
        optSource(1).Value = True
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 修改项目数据 ItemUpdate
'==============================================================================
Public Sub ItemUpdate()
    Dim strEditMode     As String
    On Error GoTo ErrH

    mEditMode = 修改
    unWinLock True
    txtName.SetFocus
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 取消保存数据 ItemDelete
'==============================================================================
Public Sub ItemDelete()
    Dim varPos  As Variant
    Dim strEditMode As String
    On Error GoTo ErrH
    
    If zlCheck.Msg_OKC("确认删除如下审查项目吗？" & vbCrLf & "编码：【" & mstrItemCode & "】" & vbCrLf & "名称：【" & mstrItemName & "】") Then Exit Sub
    AuditItemItemDelete
    mEditMode = 浏览
    mblnDataChange = False
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 保存分类数据 ItemSave
'==============================================================================
Public Function ItemSave() As Boolean
    Dim strMsg      As String
    Dim varPos      As Variant
    Dim ctrSetF     As Control
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim strSQL      As String
    Dim rsTemp      As ADODB.Recordset
    On Error GoTo ErrH
    ItemSave = True
    
    If txtTypeID.Tag <> "" Then mlngItemTypeID = Val(txtTypeID.Tag)
    mstrItemCode = txtCode.Text
    mstrItemName = txtName.Text
    mstrItemMnemonicCode = txtMnemonicCode.Text
    mstrItemDescription = txtDescription.Text
    mstrItemAudit = txtAudit_NotCheck.Text
    mintItemUsed = Val(zlCheck.Cmb_ID(cboUsed))
    mstrItemLink = Val(zlCheck.Cmb_ID(cboLink))
    mintItemFileID = txtFileID.Text
    mstrNumValue = txtNumValue.Text
    mstrPalValue = CboPalValue.ItemData(CboPalValue.ListIndex)
    mintSource = IIf(optSource(0).Value, 0, 1)
   
    If Not CheckAuditSql_IN(txtAudit_NotCheck.Text, False, IIf(optSource(0).Value, 0, 1)) Then Exit Function
    strMsg = ""
    strMsg = zlCheck.Chk_CheckTxtNull("编码", txtCode, ctrSetF, strMsg)
    '检测编码重复
    strSQL = "select count(*) from 病案审查目录 where 编码 = [1] and ID != [2]"
    If zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrItemCode, mlngItemID).Fields(0) <> "0" Then
        If ctrSetF Is Nothing Then Set ctrSetF = txtCode
        strMsg = strMsg & "编码【" & txtCode.Text & "】已存在，请重新录入或修改！" & vbCrLf
    End If
    strMsg = zlCheck.Chk_CheckTxtNull("名称", txtName, ctrSetF, strMsg)
    '检测编码重复
    If mEditMode = 新增 Or mEditMode = 复制新增 Then
        strSQL = "select count(*) from 病案审查目录 where 名称 = [1] and 分类ID = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrItemName, mlngItemTypeID, mlngItemID)
    ElseIf mEditMode = 修改 Then
        strSQL = "select count(*) from 病案审查目录 where 名称 = [1] and 分类ID = [2] and ID != [3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrItemName, mlngItemTypeID, mlngItemID)
    End If
    If rsTemp.Fields(0) <> "0" Then
        If ctrSetF Is Nothing Then Set ctrSetF = txtName
        strMsg = strMsg & "名称【" & txtName.Text & "】在分类【" & txtTypeID.Text & "】已存在，请重新录入或修改！" & vbCrLf
    End If
    
    strMsg = zlCheck.Chk_CheckTxtNull("适用对象", cboUsed, ctrSetF, strMsg)
    DoEvents
    If zlCheck.Chk_CheckMsg(strMsg, ctrSetF) Then Exit Function
    If zlCheck.Cmb_ID(cboUsed) = "" Then
        zlCheck.Msg_OK "请选择适用对象！"
        Exit Function
    End If

    Select Case mEditMode
        Case 编辑模式.新增, 编辑模式.复制新增
            AuditItemItemInsert
        Case 编辑模式.修改
            AuditItemItemUpdate
    End Select
    ItemSave = False
    mEditMode = 浏览
    mblnDataChange = False
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 取消保存数据 ItemCancel
'==============================================================================
Public Function ItemCancel() As Boolean
   
    On Error GoTo ErrH
    ItemCancel = True
    If mblnDataChange Then
        If zlCheck.Msg_OKC("确认取消你所做操作吗？") Then Exit Function
    End If
    mEditMode = 浏览
    WinLock
    mblnDataChange = False
    ItemCancel = False
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboUsed_Validate(Cancel As Boolean)
    If mEditMode <> 浏览 Then mblnDataChange = True
End Sub

Private Sub txtAudit_Change()
    If mEditMode <> 浏览 Then mblnDataChange = True
End Sub

Private Sub txtCode_Change()
    If mEditMode <> 浏览 Then mblnDataChange = True
End Sub

Private Sub txtDescription_Change()
    If mEditMode <> 浏览 Then mblnDataChange = True
End Sub

Private Sub txtMnemonicCode_Change()
    If mEditMode <> 浏览 Then mblnDataChange = True
End Sub

Private Sub txtTypeID_Change()
    If mEditMode <> 浏览 Then mblnDataChange = True
End Sub

Private Sub cboUsed_Change()
    If mEditMode <> 浏览 Then mblnDataChange = True
End Sub

Private Sub cboUsed_GotFocus()
    On Error GoTo ErrH
    mintUsedListIndex = cboUsed.ListIndex
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboUsed_LostFocus()
    On Error GoTo ErrH
    If mintUsedListIndex <> cboUsed.ListIndex Then
        txtFileID.Text = ""
        txtFileID.Tag = ""
        vsfFiles.Rows = 1
    End If
'    gstrSQL = "Select Count(*) From 病历文件列表 where 种类 = [1] "
    cmdSelectFile.Enabled = Not (Val(cboUsed.Text) = 1 Or Val(cboUsed.Text) = 5 Or Val(cboUsed.Text) = 9 Or cboUsed.Text = "")
    txtFileID.Enabled = cmdSelectFile.Enabled
    cmdSelectFile.Enabled = cmdSelectFile.Enabled
    lbl(6).Enabled = cmdSelectFile.Enabled
    vsfFiles.Enabled = cmdSelectFile.Enabled
    If vsfFiles.Enabled = True Then
        vsfFiles.BackColorBkg = &H80000005
    Else
        vsfFiles.BackColorBkg = &H8000000F
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtName_Change()
    On Error GoTo ErrH
    txtMnemonicCode.Text = zlCommFun.SpellCode(txtName.Text)
    If mEditMode <> 浏览 Then mblnDataChange = True
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
