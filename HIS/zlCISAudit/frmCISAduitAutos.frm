VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISAduitAutos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自动审查"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   Icon            =   "frmCISAduitAutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8325
      TabIndex        =   25
      Top             =   7155
      Width           =   1100
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   2
      Left            =   90
      ScaleHeight     =   5175
      ScaleWidth      =   2880
      TabIndex        =   18
      Top             =   1125
      Width           =   2880
      Begin MSComctlLib.TreeView tvw 
         Height          =   5145
         Left            =   15
         TabIndex        =   19
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
      Picture         =   "frmCISAduitAutos.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "全清"
      Top             =   6360
      Width           =   570
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "全选"
      Height          =   495
      Left            =   1755
      Picture         =   "frmCISAduitAutos.frx":00E0
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "全选"
      Top             =   6360
      Width           =   570
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9570
      TabIndex        =   1
      Top             =   7155
      Width           =   1100
   End
   Begin VB.Frame fraDetail 
      Caption         =   "基础资料"
      Height          =   1020
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
         TabIndex        =   11
         Top             =   690
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
         TabIndex        =   5
         Top             =   690
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
         TabIndex        =   4
         Top             =   690
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
         TabIndex        =   3
         Top             =   690
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
         TabIndex        =   2
         Top             =   690
         Width           =   1410
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   2265
         TabIndex        =   24
         Top             =   308
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Max             =   1000
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1/100"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   2
         Left            =   1770
         TabIndex        =   23
         Top             =   330
         Width           =   450
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人进度"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   1
         Left            =   645
         TabIndex        =   22
         Top             =   330
         Width           =   720
      End
      Begin VB.Line Line2 
         X1              =   6855
         X2              =   7140
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line1 
         X1              =   8760
         X2              =   10200
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line5 
         X1              =   4815
         X2              =   6255
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line4 
         X1              =   2745
         X2              =   4185
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line3 
         X1              =   630
         X2              =   2070
         Y1              =   885
         Y2              =   885
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
         TabIndex        =   10
         Top             =   720
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
         TabIndex        =   9
         Top             =   720
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
         TabIndex        =   8
         Top             =   720
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
         TabIndex        =   7
         Top             =   720
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
         TabIndex        =   6
         Top             =   720
         Width           =   540
      End
   End
   Begin MSComctlLib.ProgressBar pbrBar 
      Height          =   345
      Left            =   2250
      TabIndex        =   12
      Top             =   7155
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "自动(&A)"
      Height          =   350
      Left            =   7035
      TabIndex        =   13
      Top             =   7155
      Width           =   1100
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "终止(&S)"
      Height          =   350
      Left            =   7035
      TabIndex        =   14
      Top             =   7155
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFeedback 
      Height          =   5730
      Left            =   3030
      TabIndex        =   20
      Top             =   1125
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
      TabIndex        =   21
      Top             =   6375
      Width           =   1275
   End
   Begin VB.Shape shpStatus 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   480
      Left            =   90
      Top             =   6367
      Width           =   1275
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   -210
      X2              =   10770
      Y1              =   6930
      Y2              =   6945
   End
   Begin VB.Line Line6 
      X1              =   -225
      X2              =   10770
      Y1              =   7050
      Y2              =   7050
   End
   Begin VB.Label LabStatus 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   150
      TabIndex        =   15
      Top             =   7230
      Visible         =   0   'False
      Width           =   2025
   End
End
Attribute VB_Name = "frmCISAduitAutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnStop                As Boolean          '自动时停止
Private mintType                As Integer          '2抽查 1审查
Private mblnOK                  As Boolean          '确定　取消
Private mstrSortID              As String
Private mvsList                 As VSFlexGrid
Private mselectKind             As String           '选中的类型
Private mlngRows                 As Long
Public Function ShowMe(ByVal frmPar As Object, ByVal intType As Integer, ByVal vsList As VSFlexGrid) As Boolean
'2抽查 1审查
Dim i As Integer
    mintType = intType
    Set mvsList = vsList
    If mintType = 1 Then
        labShow.Caption = "审查"
        labShow.ForeColor = vbBlue
    Else
        labShow.Caption = "抽查"
        labShow.ForeColor = vbBlack
    End If
    
    If mintType = 1 Then '在院审查
        mlngRows = vsList.Rows - 1
    Else                '出院抽查
        mlngRows = 0
        With mvsList
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) > 10 Then
                mlngRows = mlngRows + 1
            End If
        Next
        End With
        If mlngRows = 0 Then
            MsgBox "当前主界面病案列表中没有可用于自动审查的病人病案，请通过“过滤”筛选或选择病人病案进行“抽审”后再试！", vbInformation, gstrSysName
            GoTo Out
        End If
    End If
    
    lblInfo(2).Caption = "0/" & mlngRows
    ProgressBar1.Max = mlngRows
    ProgressBar1.Visible = True: ProgressBar1.Value = 0
    
    Call InitVsflexGrid
    Call InitTreeView(0, 0, 0)
    'RestoreWinState Me, App.ProductName
    
    Me.Show vbModal, frmPar
Out:    Set mvsList = Nothing
        ShowMe = mblnOK
End Function
Private Sub cmdAll_Click()
    Call AllNot
End Sub
Private Sub CmdNot_Click()
    Call AllNot(False)
End Sub
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

Private Sub cmdOk_Click()
Dim i As Long
Dim lng病人ID As Long, lng主页ID As Long, lng提交Id As Long, lng反馈ID As Long, str文件id As String
Dim str意见 As String, str反馈人 As String, str反馈时间 As String, str处理期限 As String, lng科室ID As Long, str子文档ID As String

    On Error GoTo ErrH
    If vsfFeedback.Rows <= 1 Then
        MsgBox "没有生成数据，请点击取消！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With gcnOracle
        .BeginTrans
        With vsfFeedback
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("选择")) Then
                    lng病人ID = .TextMatrix(i, .ColIndex("病人ID"))
                    lng主页ID = .TextMatrix(i, .ColIndex("主页ID"))
                    lng提交Id = .TextMatrix(i, .ColIndex("提交ID"))
                    lng反馈ID = zlDatabase.GetNextId("病案反馈记录")
                    str文件id = .TextMatrix(i, .ColIndex("文件ID"))
                    str意见 = .TextMatrix(i, .ColIndex("反馈意见"))
                    str反馈人 = .TextMatrix(i, .ColIndex("反馈人"))
                    str反馈时间 = .TextMatrix(i, .ColIndex("反馈时间"))
                    str处理期限 = .TextMatrix(i, .ColIndex("处理期限"))
                    lng科室ID = .TextMatrix(i, .ColIndex("科室ID"))
                    str子文档ID = .TextMatrix(i, .ColIndex("子文档ID"))
                    
                    gstrSQL = "zl_病案反馈记录_Update (" & lng反馈ID & ",Null," & IIf(lng提交Id <= 0, "Null", lng提交Id) & "," & lng病人ID & "," & _
                              "" & lng主页ID & "," & AppObject(.TextMatrix(i, .ColIndex("反馈对象")), False) & ",'" & str文件id & "','" & str意见 & "'," & _
                              "" & .TextMatrix(i, .ColIndex("反馈项目ID")) & ",'" & str反馈人 & "',to_date('" & str反馈时间 & "','yyyy-mm-dd hh24:mi:ss'),to_date('" & str处理期限 & "','yyyy-mm-dd hh24:mi:ss')," & _
                              "" & "Null," & lng科室ID & ",null,null,null,null,null,null,'" & str子文档ID & "')"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Name
                End If
            Next
        End With
        .CommitTrans
    End With
    mblnOK = True
    Unload Me
    Exit Sub
ErrH:
    If gcnOracle.Errors.count > 0 Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdStop_Click()
    mblnStop = True
End Sub
Private Sub CheckSignle(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng提交Id As Long, ByVal datFeed As Date, ByVal datFeedBack As Date, ByVal strKey As String, ByRef blnStop As Boolean)
Dim i   As Integer, j   As Integer, str反馈人 As String, str姓名 As String
Dim varSplit        As Variant, varTmp         As Variant
Dim rsTmp   As ADODB.Recordset, rsFeed  As ADODB.Recordset
Dim strSource      As String, varPar() As String
Dim strDocid As String, strSubDocid As String, strReturn As String, strMid As String, strAlidin As String

    On Error GoTo ErrH
    str反馈人 = UserInfo.姓名
    str姓名 = txt姓名.Text
    blnStop = False
    pbrBar.Max = tvw.Nodes.count
    varSplit = Split(strKey, strSplitCmb)
    
    strKey = ""
    LabStatus.Caption = "正在分析适应对象..."
    LabStatus.BackColor = vbYellow
    DoEvents
    Sleep 200
    
    If UBound(varSplit) > 1 Then pbrBar.Max = UBound(varSplit) - 1
    For i = 0 To UBound(varSplit) - 1
        pbrBar.Value = i
        DoEvents
        If mblnStop Then blnStop = True: Exit Sub
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
                gstrSQL = "select 文件Id from 电子病历记录 where Id = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, Val(varTmp(1)))
                If Not rsTmp.EOF Then
                    strKey = strKey & Left(varSplit(i), 2) & "_" & rsTmp.Fields(0) & "[O]" & Mid(varSplit(i), 2, 1) & "[F]" & varTmp(1) & "[D],"
                End If
            End If
        ElseIf InStr(varSplit(i), "R") = 0 Then
            If Not gobjEmr Is Nothing Then
                If InStr(tvw.Nodes(varSplit(i)).Tag, "|") = 0 Then
                    strDocid = varSplit(i)
                    strSubDocid = ""
                Else
                    strDocid = Split(tvw.Nodes(varSplit(i)).Tag, "|")(0)
                    strSubDocid = Split(tvw.Nodes(varSplit(i)).Tag, "|")(1)
                End If
                gstrSQL = "Select RawtoHex(Antetype_id) as ID From bz_doc_Tasks Where Real_Doc_Id = Hextoraw(:rdid)" & IIf(strSubDocid = "", "", " And subdoc_id=:sdid")
                strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strDocid & "^" & DbType.T_String & "^rdid" & IIf(strSubDocid = "", "", "|" & strSubDocid & "^" & DbType.T_String & "^sdid"), rsTmp)
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
    strSource = "" & vbNewLine & _
                "Select x.Id, x.分类id, x.编码, x.名称, x.简码, x.说明, x.适用对象, x.适用环节, x.审查依据, '' As 文件id" & vbNewLine & _
                "From 病案审查目录 X, 病案审查方案 C, 病案审查分类 B" & vbNewLine & _
                "Where Nvl(文件id, '') Is Null And b.方案id = c.Id And b.Id = x.分类id And c.启用时间 Is Not Null And x.适用环节 = 0 Or x.适用环节 = [1]" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select x.Id, x.分类id, x.编码, x.名称, x.简码, x.说明, x.适用对象, x.适用环节, x.审查依据, y.Column_Value As 文件id" & vbNewLine & _
                "From 病案审查目录 X, 病案审查方案 C, 病案审查分类 B, Table(Cast(f_Str2list(x.文件id) As Zltools.t_Strlist)) Y" & vbNewLine & _
                "Where (x.适用环节 = 0 Or x.适用环节 = [1]) And b.方案id = c.Id And b.Id = x.分类id And c.启用时间 Is Not Null"

    gstrSQL = "" & _
            "Select a.Id,a.审查依据,b.适用对象,Decode(Length(b.文件id), 65, Substr(b.文件id, 1, 32), b.文件id) As 文件id,b.部门Id,Decode(Length(b.文件id), 65, Substr(b.文件id, 34), b.文件id) As 子文档id from (" & strSource & ") a," & vbCrLf & _
            "(" & vbCrLf & _
            "   Select" & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,1,INSTR(COLUMN_VALUE,'[O]')-1) AS Id," & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,INSTR(COLUMN_VALUE,'[O]')+length('[O]'),case when (INSTR(COLUMN_VALUE,'[F]'))-(INSTR(COLUMN_VALUE,'[O]')+length('[O]'))<0 then 1000 else (INSTR(COLUMN_VALUE,'[F]'))-(INSTR(COLUMN_VALUE,'[O]')+length('[O]')) end) as 适用对象," & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,INSTR(COLUMN_VALUE,'[F]')+length('[F]'),case when (INSTR(COLUMN_VALUE,'[D]'))-(INSTR(COLUMN_VALUE,'[F]')+length('[F]'))<0 then 1000 else (INSTR(COLUMN_VALUE,'[D]'))-(INSTR(COLUMN_VALUE,'[F]')+length('[F]')) end) as 文件Id," & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,INSTR(COLUMN_VALUE,'[D]')+length('[D]')) AS 部门Id" & vbCrLf & _
            "   From " & LongIDsTable(strKey, varPar, 2) & vbCrLf & _
            ")b" & vbCrLf & _
            "Where 'R' || to_char(a.适用对象) || Case When nvl(a.文件Id,'0')='0' Then '' Else '_' || a.文件Id End = b.Id And a.审查依据 is not null"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, CStr(mintType), varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9))
    i = 0
    LabStatus.Caption = "正在生成反馈信息..."
    LabStatus.BackColor = vbGreen
    DoEvents
    Sleep 200
    
    If Not rsTmp.EOF Then
        pbrBar.Max = rsTmp.RecordCount
    End If
    
    Do Until rsTmp.EOF
        pbrBar.Value = Val(rsTmp.Bookmark)
        DoEvents
        If mblnStop Then blnStop = True: Exit Sub
        With vsfFeedback
            If Len(NVL(rsTmp!文件ID)) < 32 Or InStr(NVL(rsTmp!ID), "R") > 0 Then
                gstrSQL = CheckAuditSql_OUT(rsTmp!审查依据, lng病人ID, lng主页ID)
                Set rsFeed = zlDatabase.OpenSQLRecord("select ZL_FUN_ExecSql('" & Replace(gstrSQL, "'", "''") & "') from dual", "mdlCISAudit")
            ElseIf Not gobjEmr Is Nothing Then
                If strMid = "" Then Call GetEMR_MID_ALIDIN(lng病人ID, lng主页ID, strMid, strAlidin) '取新病历主体ID,活动ID
                gstrSQL = Replace(rsTmp!审查依据, "[MID]", ":mid")
                gstrSQL = Replace(gstrSQL, "[ALIDIN]", ":alidin")
                strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, IIf(strMid = "", "", strMid & "^" & DbType.T_String & "^mid") & IIf(strAlidin = "", "", IIf(strMid = "", "", "|") & strAlidin & "^" & DbType.T_String & "^alidin"), rsFeed)
                If strReturn <> "" Then Set rsFeed = New ADODB.Recordset
            End If
            
            If Not rsFeed.EOF Then
            If InStr(1, rsFeed.Fields(0), "[zlsoft]Error[zlsoft]") = 0 Then
                If Trim("" & rsFeed.Fields(0)) <> "" Then
                    .Rows = .Rows + 1
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("选择")) = flexAlignCenterCenter
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("反馈意见")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("反馈对象")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("文件Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("科室Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("病人Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("主页Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("反馈人")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("反馈时间")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("处理期限")) = flexAlignLeftCenter
                    
                    .TextMatrix(.Rows - 1, .ColIndex("选择")) = True
                    .TextMatrix(.Rows - 1, .ColIndex("反馈意见")) = "" & rsFeed.Fields(0)
                    .TextMatrix(.Rows - 1, .ColIndex("反馈项目ID")) = 0 & rsTmp.Fields("Id")
                    .TextMatrix(.Rows - 1, .ColIndex("反馈对象")) = AppObject(rsTmp.Fields("适用对象"), True)
                    .TextMatrix(.Rows - 1, .ColIndex("文件Id")) = NVL(rsTmp.Fields("文件Id"))
                    .TextMatrix(.Rows - 1, .ColIndex("科室Id")) = 0 & rsTmp.Fields("部门Id")
                    .TextMatrix(.Rows - 1, .ColIndex("病人Id")) = lng病人ID
                    .TextMatrix(.Rows - 1, .ColIndex("主页Id")) = lng主页ID
                    .TextMatrix(.Rows - 1, .ColIndex("提交ID")) = lng提交Id
                    .TextMatrix(.Rows - 1, .ColIndex("姓名")) = str姓名
                    .TextMatrix(.Rows - 1, .ColIndex("Id")) = "" & .Rows - 1
                    .TextMatrix(.Rows - 1, .ColIndex("记录性质")) = 1
                    .TextMatrix(.Rows - 1, .ColIndex("记录状态")) = 1
                    .TextMatrix(.Rows - 1, .ColIndex("反馈人")) = str反馈人
                    .TextMatrix(.Rows - 1, .ColIndex("反馈时间")) = datFeed
                    .TextMatrix(.Rows - 1, .ColIndex("处理期限")) = datFeedBack
                    .TextMatrix(.Rows - 1, .ColIndex("子文档ID")) = NVL(rsTmp.Fields("子文档Id"))
                End If
            End If
            End If
        End With
        rsTmp.MoveNext
    Loop
    Exit Sub
ErrH:
    Err.Clear
End Sub
Private Sub Normal(Optional blnStar As Boolean)
    mblnStop = False
    If blnStar Then
        LabStatus.Caption = "启动自动评审"
        cmdAuto.Visible = False
        cmdStop.Visible = True
        cmdOK.Enabled = False
        cmdCancel.Enabled = False
        pbrBar.Visible = True
        LabStatus.Visible = True
        Call InitVsflexGrid
    Else
        cmdAuto.Visible = True
        cmdCancel.Enabled = True
        cmdOK.Enabled = True
        cmdStop.Visible = False
        pbrBar.Visible = False
        LabStatus.Visible = False
        Call InitTreeView(0, 0, 0)
        
        Dim i As Integer
        For i = 1 To tvw.Nodes.count
            If InStr(mselectKind, tvw.Nodes.Item(i).Key) > 0 Then
                tvw.Nodes.Item(i).Checked = True
            End If
        Next
    End If
End Sub
Private Function GetSelectKey() As String
'--检查选择分类'返回分类下所有ID
Dim i As Integer
    For i = 1 To tvw.Nodes.count
        If InStr(mselectKind, tvw.Nodes.Item(i).Key) > 0 Then
            tvw.Nodes.Item(i).Checked = True
            Call tvw_NodeCheck(tvw.Nodes.Item(i))
        End If
    Next
    
    For i = 1 To tvw.Nodes.count
        If tvw.Nodes(i).Checked Then
            GetSelectKey = GetSelectKey & tvw.Nodes.Item(i).Key & strSplitCmb
        End If
    Next
End Function
Private Function ValidateSelect() As Boolean
Dim i As Integer
    mselectKind = ""
    For i = 1 To tvw.Nodes.count
        If tvw.Nodes(i).Checked Then
            ValidateSelect = True
            mselectKind = mselectKind & tvw.Nodes.Item(i).Key & strSplitCmb
        End If
    Next
End Function
Private Sub cmdAuto_Click()
Dim i As Integer, lng病人ID As Long, lng主页ID As Long, lng科室ID As Long, lng提交Id As Long
Dim datFeed As Date, datFeedBack As Date, strKey As String, blnStop As Boolean
    
    If Not ValidateSelect Then
        MsgBox "请选择需要进行自动审查的适用对象！", vbInformation, gstrSysName: Exit Sub
    End If
    
    Call Normal(True)
    
    datFeed = zlDatabase.Currentdate
    i = Val(GetPara("反馈处理期限", 1560))
    datFeedBack = DateAdd("D", i, datFeed)
    If cmdStop.Enabled And cmdStop.Visible Then
        Call cmdStop.SetFocus
    End If
    
    With mvsList
        For i = 1 To .Rows - 1
            If mintType = 1 Or (Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) > 10) Then
                If mblnStop = True Then
                    Call Normal
                    Call MsgBox("病案自动评审中途取消，已完成部分评审！", vbCritical, gstrSysName)
                    Exit Sub
                End If
                
                lng病人ID = .TextMatrix(i, .ColIndex("病人ID"))
                lng主页ID = .TextMatrix(i, .ColIndex("主页ID"))
                lng科室ID = .TextMatrix(i, .ColIndex("出院科室ID"))
                
                If mintType = 2 Then '出院抽查才有提交ID 2-抽查　1-审查
                    lng提交Id = .TextMatrix(i, .ColIndex("ID"))
                Else
                    lng提交Id = -1
                End If
                
                Call ReadPartentInfo(lng病人ID)
                Call InitTreeView(lng病人ID, lng主页ID, lng科室ID)
                strKey = GetSelectKey
                Call CheckSignle(lng病人ID, lng主页ID, lng提交Id, datFeed, datFeedBack, strKey, blnStop)
                If blnStop = True Then
                    Call Normal
                    Call MsgBox("病案自动评审中途取消，已完成部分评审！", vbCritical, gstrSysName)
                    Exit Sub
                End If
                ProgressBar1.Value = i
                lblInfo(2).Caption = i & "/" & mlngRows
                DoEvents
                Sleep 200
            End If
        Next
    End With
    
    Call Normal
    LabStatus.Caption = "病案自动评审完成，共" & vsfFeedback.Rows - 1 & "行反馈记录尚未保存，请检查后点击<确定>保存！": LabStatus.Visible = True
End Sub
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub
Private Sub InitTreeView(ByVal lng病人ID As Long, ByVal lng主页ID As Long, lng科室ID As Long)
Dim objNode     As Node, rsTemp As ADODB.Recordset
Dim strIcon     As String, strKey     As String
Dim strSQL As String
Dim blnOldData As Boolean, strTemp As String

    On Error GoTo ErrH
    
    Set tvw.ImageList = frmPubResource.ils16
        
    If Not (tvw.SelectedItem Is Nothing) Then strKey = tvw.SelectedItem.Key
    If InStr(strKey, "K") = 0 And strKey <> "R1" And strKey <> "R5" Then strKey = ""
    
    LockWindowUpdate tvw.hWnd
    
    tvw.Nodes.Clear
    DoEvents
    
    strSQL = "Select 1 From 病人护理记录 A Where a.病人id = [1] And a.主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否存在老板数据", lng病人ID, lng主页ID)
    blnOldData = IIf(rsTemp.RecordCount > 0, True, False)
    Set rsTemp = gclsPackage.GetCISStruct(lng病人ID, lng主页ID, lng科室ID, False)
    
    Do Until rsTemp.EOF
        strIcon = zlCommFun.NVL(rsTemp("图标").Value)
        
        If zlCommFun.NVL(rsTemp("上级Id").Value) = "" Then
            Set objNode = tvw.Nodes.Add(, , rsTemp("Id").Value, rsTemp("名称").Value, strIcon, strIcon)
            objNode.Tag = zlCommFun.NVL(rsTemp("参数").Value)
        Else
            If rsTemp("上级ID").Value = "R4" Then
                strTemp = IIf(blnOldData, rsTemp("Id").Value, rsTemp("EPRID").Value)
            Else
                strTemp = rsTemp("Id").Value
            End If
            Set objNode = tvw.Nodes.Add(rsTemp("上级Id").Value, tvwChild, strTemp, rsTemp("名称").Value, strIcon, strIcon)
            objNode.Tag = zlCommFun.NVL(rsTemp("参数").Value)
        End If
        
        rsTemp.MoveNext
    Loop
    
    Set rsTemp = New ADODB.Recordset '新版病历
    Set rsTemp = gclsPackage.GetEmrCISStruct(lng病人ID, lng主页ID)
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
        
    LockWindowUpdate 0
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub ReadPartentInfo(ByVal lng病人ID As Long)
Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrH
    
    gstrSQL = "Select 住院号,住院次数,姓名,性别,年龄 From 病人信息 Where 病人Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, lng病人ID)
    If Not rsTemp.EOF Then
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

Private Sub InitVsflexGrid()
Dim strField As String, strFieldWidth  As String, varField As Variant, varFieldWidth  As Variant, i As Integer
Dim rsTemp As New ADODB.Recordset

    On Error GoTo ErrH
    
    vsfFeedback.FocusRect = flexFocusNone
    vsfFeedback.ExtendLastCol = True
    vsfFeedback.ExplorerBar = flexExSortShowAndMove
    vsfFeedback.AutoResize = False
    vsfFeedback.Editable = flexEDKbdMouse
    
    gstrSQL = "Select /*+ rule */" & vbNewLine & _
            " 0 As ID, '' As 选择,'' AS 姓名,相关id, 提交id, 病人id, 主页id, 反馈对象, 文件id, 医嘱id, 科室id, 记录性质, 记录状态, 反馈人, 反馈时间, 处理期限, 反馈意见, 反馈项目id, 子文档id" & vbNewLine & _
            "From 病案反馈记录" & vbNewLine & _
            "Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, -1)
    Set vsfFeedback.DataSource = rsTemp
    With vsfFeedback
        .FrozenCols = 3
        .ColWidth(.ColIndex("反馈时间")) = 1000
        .ColWidth(.ColIndex("处理期限")) = 1000
        .ColWidth(.ColIndex("姓名")) = 1200
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Form_Unload(Cancel)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdCancel.Enabled = False Then
        If MsgBox("需要终止已启动的自动评审吗？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call cmdStop_Click
        End If
        Cancel = -1
    End If
End Sub
Private Sub tvw_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error GoTo ErrH
    
    NoteChildChecked Node, Node.Checked
    NotePrentChecked Node, Node.Checked
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
    
Private Sub NoteChildChecked(nodex As Node, blnChecked As Boolean)
    Dim count           As Integer
    Dim ChildNode       As Node
    Dim i               As Integer
    
    On Error GoTo ErrH
    
    count = nodex.Children
    '对节点进行操作
    nodex.Checked = blnChecked
    If count > 0 Then
        Set ChildNode = nodex.Child
        NoteChildChecked ChildNode, blnChecked
        For i = 2 To count
            Set ChildNode = ChildNode.Next
            NoteChildChecked ChildNode, blnChecked
        Next
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub NotePrentChecked(nodex As Node, blnChecked As Boolean)
    On Error GoTo ErrH
    If blnChecked And (Not nodex.Parent Is Nothing) Then nodex.Parent.Checked = True
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
Private Sub vsfFeedback_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfFeedback.ColIndex("选择") Then Cancel = True
End Sub
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

