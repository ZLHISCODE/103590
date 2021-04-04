VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISAduitAuto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自动审查"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   Icon            =   "frmCISAduitAuto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   2
      Left            =   90
      ScaleHeight     =   5175
      ScaleWidth      =   2880
      TabIndex        =   19
      Top             =   780
      Width           =   2880
      Begin MSComctlLib.TreeView tvw 
         Height          =   5145
         Left            =   15
         TabIndex        =   20
         Top             =   15
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   9075
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   0
      End
   End
   Begin VB.CommandButton CmdNot 
      Caption         =   "全清"
      Height          =   495
      Left            =   2370
      Picture         =   "frmCISAduitAuto.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "全清"
      Top             =   6015
      Width           =   570
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "全选"
      Height          =   495
      Left            =   1755
      Picture         =   "frmCISAduitAuto.frx":00E0
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "全选"
      Top             =   6015
      Width           =   570
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9570
      TabIndex        =   2
      Top             =   7020
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8280
      TabIndex        =   1
      Top             =   7020
      Width           =   1100
   End
   Begin VB.Frame fraDetail 
      Caption         =   "基础资料"
      Height          =   690
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   10590
      Begin VB.TextBox txt住院号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   8775
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   12
         Top             =   315
         Width           =   1410
      End
      Begin VB.TextBox txt年龄 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   315
         Width           =   1410
      End
      Begin VB.TextBox txt姓名 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   645
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   5
         Top             =   315
         Width           =   1410
      End
      Begin VB.TextBox txt住院次数 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   6855
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   285
      End
      Begin VB.TextBox txt性别 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   3
         Top             =   315
         Width           =   1410
      End
      Begin VB.Line Line2 
         X1              =   6855
         X2              =   7140
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line1 
         X1              =   8760
         X2              =   10200
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line5 
         X1              =   4815
         X2              =   6255
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line4 
         X1              =   2745
         X2              =   4185
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line3 
         X1              =   630
         X2              =   2070
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   7
         Left            =   4320
         TabIndex        =   11
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   5
         Left            =   2265
         TabIndex        =   10
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   9
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "第    次住院"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   3
         Left            =   6645
         TabIndex        =   8
         Top             =   345
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   0
         Left            =   8145
         TabIndex        =   7
         Top             =   345
         Width           =   540
      End
   End
   Begin MSComctlLib.ProgressBar pbrBar 
      Height          =   345
      Left            =   2250
      TabIndex        =   13
      Top             =   7020
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "自动(&A)"
      Height          =   350
      Left            =   6930
      TabIndex        =   14
      Top             =   7020
      Width           =   1100
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "终止(&S)"
      Height          =   350
      Left            =   6930
      TabIndex        =   15
      Top             =   7020
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFeedback 
      Height          =   5730
      Left            =   3030
      TabIndex        =   21
      Top             =   780
      Width           =   7635
      _cx             =   13467
      _cy             =   10107
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.Label labShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "审查"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   90
      TabIndex        =   22
      Top             =   6030
      Width           =   1275
   End
   Begin VB.Shape shpStatus 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   480
      Left            =   90
      Top             =   6030
      Width           =   1275
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   -210
      X2              =   10770
      Y1              =   6720
      Y2              =   6735
   End
   Begin VB.Line Line6 
      X1              =   -225
      X2              =   10770
      Y1              =   6705
      Y2              =   6705
   End
   Begin VB.Label LabStatus 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   150
      TabIndex        =   16
      Top             =   7095
      Visible         =   0   'False
      Width           =   2025
   End
End
Attribute VB_Name = "frmCISAduitAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng提交Id              As Long             '病人Id
Private mlng病人ID              As Long             '病人Id
Private mlng主页ID              As Long             '主页Id
Private mlng科室ID              As Long             '科室Id
Private mblnCancel              As Boolean          '确定 or 取消
Private zlCheck                 As New clsCheck     '检测
Private mstrTreeSelect          As String           '树选中的Key
Private mblnStop                As Boolean          '自动时停止
Private mblnChecked             As Boolean          '选中/取消选中
Private mstrSortID              As String
Private mstrLink                As String           '1、审查 2、抽查

'查找字段
Const con_vsfField = "/*+ rule */rownum as Id,'' as 选择,相关Id,提交Id,病人Id,主页ID,反馈对象,文件Id,医嘱Id,科室Id,记录性质,记录状态,反馈人,反馈时间,处理期限,反馈意见,反馈项目ID,子文档ID"

Public Property Get blnCancel() As Boolean
    blnCancel = mblnCancel
End Property

Public Property Let blnCancel(ByVal vNewValue As Boolean)
    mblnCancel = vNewValue
End Property

Public Property Get lng提交Id() As Long
    lng提交Id = mlng提交Id
End Property

Public Property Let lng提交Id(ByVal vNewValue As Long)
    mlng提交Id = vNewValue
End Property

Public Property Get lng病人id() As Long
    lng病人id = mlng病人ID
End Property

Public Property Let lng病人id(ByVal vNewValue As Long)
    mlng病人ID = vNewValue
End Property

Public Property Get lng主页ID() As Long
    lng主页ID = mlng主页ID
End Property

Public Property Let lng主页ID(ByVal vNewValue As Long)
    mlng主页ID = vNewValue
End Property

Public Property Get lng科室ID() As Long
    lng科室ID = mlng科室ID
End Property

Public Property Let lng科室ID(ByVal vNewValue As Long)
    mlng科室ID = vNewValue
End Property

Public Property Get strTreeSelect() As String
    strTreeSelect = mstrTreeSelect
End Property

Public Property Let strTreeSelect(ByVal vNewValue As String)
    mstrTreeSelect = vNewValue
End Property

Public Property Get strLink() As String
    strLink = mstrLink
End Property

Public Property Let strLink(ByVal vNewValue As String)
    mstrLink = vNewValue
End Property

'==============================================================================
'=功能： 全选
'==============================================================================
Private Sub cmdAll_Click()
    Call AllNot
End Sub

'==============================================================================
'=功能： 全清
'==============================================================================
Private Sub CmdNot_Click()
    Call AllNot(False)
End Sub

'==============================================================================
'=功能： 全选、全清
'==============================================================================
Private Sub AllNot(Optional blnAll As Boolean = True)
    Dim i           As Long
    On Error GoTo ErrH
    For i = 1 To tvw.Nodes.count
        tvw.Nodes.Item(i).Checked = blnAll
    Next
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


'==============================================================================
'=功能： 保存数据
'==============================================================================
Private Sub CmdOK_Click()
    Dim i       As Long

    On Error GoTo ErrH
    If vsfFeedback.Rows <= 1 Then
        zlCheck.Msg_OK "没有生成数据，请点击取消！"
        Exit Sub
    End If
    gstrSQL = ""
    With gcnOracle
        .BeginTrans
        With vsfFeedback
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("选择")) Then
                    gstrSQL = "zl_病案反馈记录_Update (" & zlDatabase.GetNextId("病案反馈记录") & ",Null," & IIf(lng提交Id <= 0, "Null", lng提交Id) & "," & lng病人id & "," & _
                              "" & lng主页ID & "," & AppObject(.TextMatrix(i, .ColIndex("反馈对象")), False) & ",'" & .TextMatrix(i, .ColIndex("文件ID")) & "','" & .TextMatrix(i, .ColIndex("反馈意见")) & "'," & _
                              "" & .TextMatrix(i, .ColIndex("反馈项目ID")) & ",'" & .TextMatrix(i, .ColIndex("反馈人")) & "',to_date('" & .TextMatrix(i, .ColIndex("反馈时间")) & "','yyyy-mm-dd hh24:mi:ss'),to_date('" & .TextMatrix(i, .ColIndex("处理期限")) & "','yyyy-mm-dd hh24:mi:ss')," & _
                              "" & "Null," & .TextMatrix(i, .ColIndex("科室Id")) & ",null,null,null,null,null,null,'" & .TextMatrix(i, .ColIndex("子文档ID")) & "')"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Name
                End If
            Next
        End With
        .CommitTrans
    End With
    blnCancel = False
    Unload Me
    Exit Sub
ErrH:
    If gcnOracle.Errors.count > 0 Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then Resume
   
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 停止
'==============================================================================
Private Sub cmdStop_Click()
    mblnStop = True
End Sub

'==============================================================================
'=功能： 自动处理
'==============================================================================
Private Sub cmdAuto_Click()
Dim i   As Integer, j   As Integer
Dim strKey          As String
Dim varSplit        As Variant, varTmp         As Variant
Dim rsTmp   As ADODB.Recordset, rsFeed  As ADODB.Recordset
Dim strSql          As String
Dim datFeed         As Date, datFeedBack    As Date
Dim intRow          As Integer, strSource      As String
Dim strDocid As String, strSubDocid As String, strReturn As String, strMid As String, strAlidin As String
    
    On Error GoTo ErrH
    
    If zlCheck.Msg_OKC("确认进行自动评审吗？" & vbCrLf & "自动评审将清空表格中现有数据！") Then Exit Sub
    LockWindowUpdate vsfFeedback.hWnd
    
    cmdAuto.Visible = False
    cmdStop.Visible = True
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    pbrBar.Visible = True
    LabStatus.Visible = True
    
    
    datFeed = zlDatabase.Currentdate
    i = Val(GetPara("反馈处理期限", 1560))
    datFeedBack = DateAdd("D", i, datFeed)
    LabStatus.Caption = "正在读取适用对象..."
    LabStatus.BackColor = vbRed
    DoEvents
    Sleep (1000)
    
    pbrBar.Max = tvw.Nodes.count
    For i = 1 To tvw.Nodes.count
        pbrBar.Value = i
        DoEvents
        If mblnStop Then
            cmdAuto.Visible = True
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            LabStatus.Visible = False
            cmdStop.Visible = False
            pbrBar.Visible = False
            mblnStop = False
            Call zlCheck.Msg_OK("病案自动评审中途取消，已完成部分评审！", vbCritical)
            LockWindowUpdate 0
            Exit Sub
        End If
        If tvw.Nodes.Item(i).Checked Then
            strKey = strKey & tvw.Nodes.Item(i).Key & strSplitCmb
        End If
    Next
    If strKey = "" Then
        zlCheck.Msg_OK ("当前未选择适用对象！")
        cmdAuto.Visible = True
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            LabStatus.Visible = False
            cmdStop.Visible = False
            pbrBar.Visible = False
            mblnStop = False
        Exit Sub
    End If
    varSplit = Split(strKey, strSplitCmb)
    
    strKey = ""
    LabStatus.Caption = "正在分析适应对象..."
    LabStatus.BackColor = vbYellow
    DoEvents
    Sleep (1000)
    
    If UBound(varSplit) > 1 Then pbrBar.Max = UBound(varSplit) - 1
    For i = 0 To UBound(varSplit) - 1
        pbrBar.Value = i
        DoEvents
        If mblnStop Then
            cmdAuto.Visible = True
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            
            LabStatus.Visible = False
            cmdStop.Visible = False
            pbrBar.Visible = False
            mblnStop = False
            Call zlCheck.Msg_OK("病案自动评审中途取消，已完成部分评审！", vbCritical)
            LockWindowUpdate 0
            Exit Sub
        End If
        If Len(varSplit(i)) = 2 Then
            strKey = strKey & varSplit(i) & "[O]" & Mid(varSplit(i), 2, 1) & "[F][D],"
             
        ElseIf InStr(1, varSplit(i), "R4") > 0 Then
            '病人护理记录，直接为文件Id
            varSplit(i) = Replace(varSplit(i), "K", ",")
            varTmp = Split(varSplit(i), ",")
            If UBound(varTmp) > 1 Then
                strKey = strKey & Left(varSplit(i), 2) & "_" & varTmp(1) & "[O]" & Mid(varSplit(i), 2, 1) & "[F]" & varTmp(1) & "[D]" & varTmp(3) & ","
            End If
        ElseIf InStr("R2R3R6R7R8", Left(varSplit(i), 2)) > 0 Then
            '电子病历记录中查找Id
            varSplit(i) = Replace(varSplit(i), "K", ",")
            varTmp = Split(varSplit(i), ",")
            If UBound(varTmp) > 1 Then
                strSql = "select 文件Id from 电子病历记录 where Id = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Name, Val(varTmp(1)))
                If Not zlCheck.Connection_ChkRsState(rsTmp) Then
                    strKey = strKey & Left(varSplit(i), 2) & "_" & rsTmp.Fields(0) & "[O]" & Mid(varSplit(i), 2, 1) & "[F]" & varTmp(1) & "[D],"
                End If
            End If
        ElseIf InStr(varSplit(i), "R") = 0 Then
            If gobjEmr Is Nothing Then
                MsgBox "本机未安装病历组件，不能进行自动审查，请检查！", vbInformation, gstrSysName
                mblnStop = True
            Else
                If InStr(tvw.Nodes(varSplit(i)).Tag, "|") = 0 Then
                    strDocid = varSplit(i)
                    strSubDocid = ""
                Else
                    strDocid = Split(tvw.Nodes(varSplit(i)).Tag, "|")(0)
                    strSubDocid = Split(tvw.Nodes(varSplit(i)).Tag, "|")(1)
                End If
                strSql = "Select RawtoHex(Antetype_id) as ID From bz_doc_Tasks Where Real_Doc_Id = Hextoraw(:rdid)" & IIf(strSubDocid = "", "", " And subdoc_id=:sdid")
                strReturn = gobjEmr.OpenSQLRecordset(strSql, strDocid & "^" & DbType.T_String & "^rdid" & IIf(strSubDocid = "", "", "|" & strSubDocid & "^" & DbType.T_String & "^sdid"), rsTmp)
                If strReturn = "" Then
                If Not rsTmp.EOF Then
                    strKey = strKey & tvw.Nodes(varSplit(i)).Parent.Key & "_" & rsTmp.Fields(0) & "[O]" & Mid(tvw.Nodes(varSplit(i)).Parent.Key, 2, 1) & "[F]" & tvw.Nodes(varSplit(i)).Tag & "[D],"
                End If
                End If
            End If
        End If
    Next
    strKey = Left(strKey, Len(strKey) - 1)
    '读取检测条件
    strSource = "" & _
            "select x.id,x.分类id,x.编码,x.名称,x.简码,x.说明,x.适用对象,x.适用环节,x.审查依据,'' as 文件ID  from 病案审查目录  x,病案审查方案 C,病案审查分类 B where nvl(文件ID,'') is null And  B.方案id = C.id and B.id =x.分类ID And  C.启用时间 is not null And x.适用环节 = 0 or x.适用环节 = [2] union all " & vbCrLf & _
            "select x.id,x.分类id,x.编码,x.名称,x.简码,x.说明,x.适用对象,x.适用环节,x.审查依据,y.column_value as 文件ID  from 病案审查目录  x , 病案审查方案 C,病案审查分类 B,table (Cast(f_Str2List( x.文件ID) As zlTools.t_StrList)) y" & vbCrLf & _
            "Where (X.适用环节 = 0 Or X.适用环节 = [2])  and  B.方案id = C.id and B.id =X.分类ID And  C.启用时间 is not null "
    
    strSql = "" & _
            "Select a.Id,a.审查依据,b.适用对象,Decode(Length(b.文件id), 65, Substr(b.文件id, 1, 32), b.文件id) As 文件id,b.部门Id,Decode(Length(b.文件id), 65, Substr(b.文件id, 34), b.文件id) As 子文档id from (" & strSource & ") a," & vbCrLf & _
            "(" & vbCrLf & _
            "   Select" & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,1,INSTR(COLUMN_VALUE,'[O]')-1) AS Id," & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,INSTR(COLUMN_VALUE,'[O]')+length('[O]'),case when (INSTR(COLUMN_VALUE,'[F]'))-(INSTR(COLUMN_VALUE,'[O]')+length('[O]'))<0 then 1000 else (INSTR(COLUMN_VALUE,'[F]'))-(INSTR(COLUMN_VALUE,'[O]')+length('[O]')) end) as 适用对象," & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,INSTR(COLUMN_VALUE,'[F]')+length('[F]'),case when (INSTR(COLUMN_VALUE,'[D]'))-(INSTR(COLUMN_VALUE,'[F]')+length('[F]'))<0 then 1000 else (INSTR(COLUMN_VALUE,'[D]'))-(INSTR(COLUMN_VALUE,'[F]')+length('[F]')) end) as 文件Id," & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,INSTR(COLUMN_VALUE,'[D]')+length('[D]')) AS 部门Id" & vbCrLf & _
            "   From Table (Cast(f_Str2List([1]) As zlTools.t_StrList))" & vbCrLf & _
            ")b" & vbCrLf & _
            "Where 'R' || to_char(a.适用对象) || Case When nvl(a.文件Id,'0')='0' Then '' Else '_' || a.文件Id End = b.Id And a.审查依据 is not null"
    If Len(strKey) >= 4000 Then
        cmdAuto.Visible = True
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
        cmdStop.Visible = False
        pbrBar.Visible = False
        LabStatus.Visible = False
        zlCheck.Msg_OK "选择项目过多，请取消部分项目。"
        LockWindowUpdate 0
        Exit Sub
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Name, strKey, mstrLink)
    i = 0
    LabStatus.Caption = "正在生成反馈信息..."
    LabStatus.BackColor = vbGreen
    DoEvents
    Sleep (1000)
    
    If rsTmp.RecordCount > 0 Then pbrBar.Max = rsTmp.RecordCount
    
    intRow = 0
    Do While Not zlCheck.Connection_ChkRsState(rsTmp)
        pbrBar.Value = Val(rsTmp.Bookmark)
        DoEvents
        If mblnStop Then
            cmdAuto.Visible = True
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            
            LabStatus.Visible = False
            LabStatus.Visible = False
            cmdStop.Visible = False
            pbrBar.Visible = False
            mblnStop = False
            Call zlCheck.Msg_OK("病案自动评审中途取消，已完成部分评审！", vbCritical)
            LockWindowUpdate 0
            Exit Sub
        End If
        With vsfFeedback
            .Rows = rsTmp.RecordCount + 10
            If Len(NVL(rsTmp!文件ID)) < 32 Or InStr(NVL(rsTmp!ID), "R") > 0 Then
                strSql = CheckAuditSql_OUT(rsTmp!审查依据, lng病人id, lng主页ID)
                Set rsFeed = zlDatabase.OpenSQLRecord("select ZL_FUN_ExecSql('" & Replace(strSql, "'", "''") & "') from dual", "mdlCISAudit")
            ElseIf Not gobjEmr Is Nothing Then
                If strMid = "" Then Call GetEMR_MID_ALIDIN(lng病人id, lng主页ID, strMid, strAlidin) '取新病历主体ID,活动ID
                strSql = Replace(rsTmp!审查依据, "[MID]", ":mid")
                strSql = Replace(rsTmp!审查依据, "[ALIDIN]", ":alidin")
                strReturn = gobjEmr.OpenSQLRecordset(strSql, IIf(strMid = "", "", strMid & "^" & DbType.T_String & "^mid") & IIf(strAlidin = "", "", IIf(strMid = "", "", "|") & strAlidin & "^" & DbType.T_String & "^alidin"), rsFeed)
                If strReturn <> "" Then Set rsFeed = New ADODB.Recordset
            End If
            
            If Not zlCheck.Connection_ChkRsState(rsFeed) Then
            If InStr(1, rsFeed.Fields(0), "[zlsoft]Error[zlsoft]") = 0 Then
                If Trim("" & rsFeed.Fields(0)) <> "" Then

                    intRow = intRow + 1
                   
                    .Cell(flexcpAlignment, intRow, .ColIndex("选择")) = flexAlignCenterCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("反馈意见")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("反馈对象")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("文件Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("科室Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("病人Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("主页Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("反馈人")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("反馈时间")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("处理期限")) = flexAlignLeftCenter
                    
                    .TextMatrix(intRow, .ColIndex("选择")) = True
                    .TextMatrix(intRow, .ColIndex("反馈意见")) = "" & rsFeed.Fields(0)
                    .TextMatrix(intRow, .ColIndex("反馈项目ID")) = 0 & rsTmp.Fields("Id")
                    .TextMatrix(intRow, .ColIndex("反馈对象")) = AppObject(rsTmp.Fields("适用对象"), True)
                    .TextMatrix(intRow, .ColIndex("文件Id")) = NVL(rsTmp.Fields("文件Id"))
                    .TextMatrix(intRow, .ColIndex("科室Id")) = 0 & rsTmp.Fields("部门Id")
                    .TextMatrix(intRow, .ColIndex("病人Id")) = lng病人id
                    .TextMatrix(intRow, .ColIndex("主页Id")) = lng主页ID
                    .TextMatrix(intRow, .ColIndex("Id")) = "" & intRow
                    .TextMatrix(intRow, .ColIndex("记录性质")) = 1
                    .TextMatrix(intRow, .ColIndex("记录状态")) = 1
                    .TextMatrix(intRow, .ColIndex("反馈人")) = UserInfo.姓名
                    .TextMatrix(intRow, .ColIndex("反馈时间")) = datFeed
                    .TextMatrix(intRow, .ColIndex("处理期限")) = datFeedBack
                    .TextMatrix(intRow, .ColIndex("子文档ID")) = NVL(rsTmp.Fields("子文档Id"))
                End If
            End If
            End If
        End With
        rsTmp.MoveNext
    Loop
    vsfFeedback.Rows = intRow + 1
    zlCheck.Msg_OK ("病案自动评审成功！")
    cmdAuto.Visible = True
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdStop.Visible = False
    pbrBar.Visible = False
    LabStatus.Visible = False
    LockWindowUpdate 0
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Call zlCheck.Msg_OK("病案自动评审失败！", vbCritical)
    cmdAuto.Visible = True
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdStop.Visible = False
    pbrBar.Visible = False
    LabStatus.Visible = False
    LockWindowUpdate 0
End Sub

'==============================================================================
'=功能： 取消
'==============================================================================
Private Sub CmdCancel_Click()
    On Error GoTo ErrH
    mblnCancel = True
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 初始化树形列表
'==============================================================================
Private Sub InitTreeView()
Dim objNode     As Node, rsTemp As ADODB.Recordset
Dim strIcon     As String, strKey     As String
Dim strSql As String
Dim blnOldData As Boolean, strTemp As String
    On Error GoTo ErrH
    
    Set tvw.ImageList = frmPubResource.ils16
        
    If Not (tvw.SelectedItem Is Nothing) Then strKey = tvw.SelectedItem.Key
    If InStr(strKey, "K") = 0 And strKey <> "R1" And strKey <> "R5" Then strKey = ""
    
    LockWindowUpdate tvw.hWnd
    
    tvw.Nodes.Clear
    DoEvents
    
    strSql = "Select 1 From 病人护理记录 A Where a.病人id = [1] And a.主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "检查是否存在老板数据", mlng病人ID, mlng主页ID)
    blnOldData = IIf(rsTemp.RecordCount > 0, True, False)
    Set rsTemp = gclsPackage.GetCISStruct(mlng病人ID, mlng主页ID, mlng科室ID, False)
    
    Do While Not zlCheck.Connection_ChkRsState(rsTemp)
        strIcon = zlCommFun.NVL(rsTemp("图标").Value)
        
        If zlCommFun.NVL(rsTemp("上级Id").Value) = "" Then
            Set objNode = tvw.Nodes.Add(, , rsTemp("Id").Value, rsTemp("名称").Value, strIcon, strIcon)
            objNode.Tag = zlCommFun.NVL(rsTemp("参数").Value)
        Else
            If rsTemp("上级ID").Value = "R4" Then
                strTemp = IIf(blnOldData, rsTemp("Id").Value, rsTemp("EPRID").Value)
                If mstrTreeSelect = rsTemp("Id").Value Then mstrTreeSelect = IIf(blnOldData, rsTemp("Id").Value, rsTemp("EPRID").Value)
            Else
                strTemp = rsTemp("Id").Value
            End If
            Set objNode = tvw.Nodes.Add(rsTemp("上级Id").Value, tvwChild, strTemp, rsTemp("名称").Value, strIcon, strIcon)
            objNode.Tag = zlCommFun.NVL(rsTemp("参数").Value)
        End If
        
        rsTemp.MoveNext
    Loop
    
    Set rsTemp = New ADODB.Recordset '新版病历
    Set rsTemp = gclsPackage.GetEmrCISStruct(mlng病人ID, mlng主页ID)
    If Not rsTemp Is Nothing Then
    If rsTemp.State = ADODB.adStateOpen Then
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        Do Until rsTemp.EOF
            Set objNode = tvw.Nodes.Add(rsTemp!上级ID.Value, tvwChild, rsTemp!ID.Value, rsTemp!名称.Value, rsTemp!图标.Value, rsTemp!图标.Value)
            objNode.Tag = NVL(rsTemp!参数) '文档ID[|子文档ID]
            rsTemp.MoveNext
        Loop
    End If
    End If
    End If
    
    If tvw.Nodes.count > 0 Then
        strKey = mstrTreeSelect
        If strKey <> "" Then
            tvw.Nodes(strKey).Selected = True
            tvw.Nodes(strKey).Expanded = True
            tvw.Nodes(strKey).Checked = True
        End If
    End If
    
    LockWindowUpdate 0
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 病人信息部份
'==============================================================================
Private Sub InitBase()
Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrH
    
    gstrSQL = "Select 住院号,住院次数,姓名,性别,年龄 From 病人信息 Where 病人Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, mlng病人ID)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        txt住院号.Text = "" & rsTemp.Fields!住院号
        txt住院次数.Text = "" & rsTemp.Fields!住院次数
        txt姓名.Text = "" & rsTemp.Fields!姓名
        txt性别.Text = "" & rsTemp.Fields!性别
        txt年龄.Text = "" & rsTemp.Fields!年龄
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    GetPersonSet = False
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 初始化网格 VsflexGrId
'==============================================================================
Private Sub InitVsflexGrid()
Dim strField As String, strFieldWidth  As String, varField As Variant, varFieldWidth  As Variant, i As Integer
Dim rsTemp As New ADODB.Recordset

    On Error GoTo ErrH
    
    vsfFeedback.FocusRect = flexFocusNone
    vsfFeedback.ExtendLastCol = True
    vsfFeedback.ExplorerBar = flexExSortShowAndMove
    vsfFeedback.AutoResize = False
    vsfFeedback.Editable = flexEDKbdMouse
    
    gstrSQL = "Select " & con_vsfField & vbCrLf & _
                "From 病案反馈记录" & vbCrLf & _
                "Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, -1)
    Set vsfFeedback.DataSource = rsTemp
    With vsfFeedback
        .FrozenCols = 3
        .ColWidth(.ColIndex("反馈时间")) = 1000
        .ColWidth(.ColIndex("处理期限")) = 1000
        If GetPersonSet Then
            '使用个性化设置【调已保存的格式】
            strField = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name & "\VSFlexGrId", vsfFeedback.Name & "名称", "")
            strFieldWidth = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name & "\VSFlexGrId", vsfFeedback.Name & "宽度", "")
            varField = Split(strField, ",")
            varFieldWidth = Split(strFieldWidth, ",")
            For i = 0 To UBound(varField)
                If varField(i) <> "" Then
                    .ColPosition(.ColIndex(varField(i))) = i
                    .ColWidth(i) = Val(varFieldWidth(i))
                End If
            Next
        End If
        .ColWidth(0) = 250
        .ColWidth(.ColIndex("选择")) = 450
        .ColDataType(.ColIndex("选择")) = flexDTBoolean
        
        .Cell(flexcpData, 0, .ColIndex("选择")) = "[选择]"
        
        .TextMatrix(0, .ColIndex("选择")) = ""
        .Cell(flexcpPicture, 0, .ColIndex("选择")) = frmPubResource.ils16.ListImages(4).Picture
        
        .Cell(flexcpPictureAlignment, 0, .ColIndex("选择")) = flexAlignCenterCenter
        
        .MergeCol(.ColIndex("分类Id")) = True
        .ColWidth(0) = 0:  .ColHidden(0) = True
        .ColWidth(.ColIndex("Id")) = 0: .ColHidden(.ColIndex("Id")) = True
        .ColWidth(.ColIndex("相关Id")) = 0: .ColHidden(.ColIndex("相关Id")) = True
        .ColWidth(.ColIndex("提交Id")) = 0: .ColHidden(.ColIndex("提交Id")) = True
        .ColWidth(.ColIndex("病人Id")) = 0: .ColHidden(.ColIndex("病人Id")) = True
        .ColWidth(.ColIndex("主页Id")) = 0: .ColHidden(.ColIndex("主页Id")) = True
        .ColWidth(.ColIndex("文件Id")) = 0: .ColHidden(.ColIndex("文件Id")) = True
        .ColWidth(.ColIndex("医嘱Id")) = 0: .ColHidden(.ColIndex("医嘱Id")) = True
        .ColWidth(.ColIndex("科室Id")) = 0: .ColHidden(.ColIndex("科室Id")) = True
        .ColWidth(.ColIndex("记录性质")) = 0: .ColHidden(.ColIndex("记录性质")) = True
        .ColWidth(.ColIndex("记录状态")) = 0: .ColHidden(.ColIndex("记录状态")) = True
        .ColWidth(.ColIndex("反馈项目ID")) = 0: .ColHidden(.ColIndex("反馈项目ID")) = True
        .ColWidth(.ColIndex("子文档Id")) = 0: .ColHidden(.ColIndex("子文档Id")) = False
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
        Next
        '可修改列
    End With
    DoEvents
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 页面初始化
'==============================================================================
Private Sub Form_Load()
    
    On Error GoTo ErrH
    zlCheck.Sys_System Me
    
    txt住院号.BackColor = fraDetail.BackColor
    txt住院次数.BackColor = fraDetail.BackColor
    txt姓名.BackColor = fraDetail.BackColor
    txt性别.BackColor = fraDetail.BackColor
    txt年龄.BackColor = fraDetail.BackColor
    Call InitTreeView
    Call InitVsflexGrid
    Call InitBase
    If mstrLink = "1" Then
        labShow.Caption = "审查"
        labShow.ForeColor = vbBlue
    Else
        labShow.Caption = "抽查"
        labShow.ForeColor = vbBlack
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveFlexState vsfFeedback, Me.Name
    Set zlCheck = Nothing
End Sub

Private Sub tvw_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error GoTo ErrH
    mblnChecked = Node.Checked
    NoteChildChecked Node
    NotePrentChecked Node
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
    
Private Sub NoteChildChecked(nodex As Node)
    Dim count           As Integer
    Dim ChildNode       As Node
    Dim i               As Integer
    
    On Error GoTo ErrH
    
    count = nodex.Children
    '对节点进行操作
    nodex.Checked = mblnChecked
    If count > 0 Then
        Set ChildNode = nodex.Child
        NoteChildChecked ChildNode
        For i = 2 To count
            Set ChildNode = ChildNode.Next
            NoteChildChecked ChildNode
        Next
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub NotePrentChecked(nodex As Node)
    On Error GoTo ErrH
    If mblnChecked And (Not nodex.Parent Is Nothing) Then nodex.Parent.Checked = True
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AppObject(strApp As String, Optional blnApp As Boolean = True) As String
    Dim strReturn       As String
    
    On Error GoTo ErrH
    
    If blnApp Then
        Select Case strApp
            Case "1"
                strReturn = "住院医嘱"
            Case "2"
                strReturn = "住院病历"
            Case "3"
                strReturn = "护理病历"
            Case "4"
                strReturn = "护理记录"
            Case "5"
                strReturn = "首页记录"
            Case "6"
                strReturn = "医嘱报告"
            Case "7"
                strReturn = "疾病证明"
            Case "8"
                strReturn = "知情文件"
            Case "9"
                strReturn = "临床路径"
        End Select
    Else
        Select Case strApp
            Case "住院医嘱"
                strReturn = "1"
            Case "住院病历"
                strReturn = "2"
            Case "护理病历"
                strReturn = "3"
            Case "护理记录"
                strReturn = "4"
            Case "首页记录"
                strReturn = "5"
            Case "医嘱报告"
                strReturn = "6"
            Case "疾病证明"
                strReturn = "7"
            Case "知情文件"
                strReturn = "8"
            Case "临床路径"
                strReturn = "9"
        End Select
    End If
    AppObject = strReturn
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsfFeedback_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrH
    vsfFeedback.TextMatrix(Row, Col) = ConvertString(vsfFeedback.TextMatrix(Row, Col))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo ErrH
    With vsfFeedback
        Select Case Col
            Case .ColIndex("反馈意见")
                vsfFeedback.ComboList = "|..."
            Case .ColIndex("选择")
                .ComboList = ""
            Case Else
                .ComboList = ""
                Cancel = True
        End Select
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 排序后定位记录 vsfFeedback
'==============================================================================
Private Sub vsfFeedback_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngRow      As Long
    On Error GoTo ErrH
    lngRow = vsfFeedback.FindRow(mstrSortID, -1, vsfFeedback.ColIndex("ID"), False, True)
    If lngRow > 0 Then vsfFeedback.Row = lngRow
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col = vsfFeedback.ColIndex("选择") Then
        Position = -1
    Else
        If Position <= vsfFeedback.ColIndex("选择") Then Position = Col
    End If
End Sub

'==============================================================================
'=功能： 某列不能拖动大小 vsfAuditItem[图标]
'==============================================================================
Private Sub vsfFeedback_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfFeedback.ColIndex("选择") Then Cancel = True
End Sub

'==============================================================================
'=功能： 排序前记录ID vsfFeedback
'==============================================================================
Private Sub vsfFeedback_BeforeSort(ByVal Col As Long, Order As Integer)
    Dim i           As Long
    Dim blnCheck    As Boolean
    On Error GoTo ErrH
    If Col = vsfFeedback.ColIndex("选择") Then
        Order = -1
        With vsfFeedback
            If .Rows <= 1 Then Exit Sub
            blnCheck = Not (.TextMatrix(1, .ColIndex("选择")) = "True")
            If blnCheck Then
                .Cell(flexcpPicture, 0, .ColIndex("选择")) = frmPubResource.ils16.ListImages(4).Picture
            Else
                .Cell(flexcpPicture, 0, .ColIndex("选择")) = frmPubResource.ils16.ListImages(25).Picture
            End If
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("选择")) = blnCheck
            Next
        End With
    End If
    mstrSortID = "" & vsfFeedback.TextMatrix(vsfFeedback.Row, vsfFeedback.ColIndex("ID"))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrH
    If vsfFeedback.ColIndex("反馈意见") = Col Then
        vsfFeedback.TextMatrix(Row, Col) = Big_Note(vsfFeedback.TextMatrix(Row, Col), vsfFeedback.ColKey(Col) & "—编辑窗口", False)
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If vsfFeedback.ColIndex("反馈意见") = vsfFeedback.Col Then
        '空格编辑
        If KeyAscii = vbKeySpace Then
            'KeyAscii = 39
            KeyAscii = 0
            SendKeys "{f2}"
        End If
        '回车 下一条编辑
        If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{down}"
            SendKeys "{f2}"
        End If
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = Asc("'") Then
       KeyAscii = 0
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


