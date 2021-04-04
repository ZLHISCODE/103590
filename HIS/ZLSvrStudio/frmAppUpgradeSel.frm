VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppUpgradeSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "历史库选择"
   ClientHeight    =   6780
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   11280
   Icon            =   "frmAppUpgradeSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   9720
      TabIndex        =   9
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdNotSel 
      Caption         =   "全消(&R)"
      Height          =   350
      Left            =   8520
      TabIndex        =   6
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelALl 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   7440
      TabIndex        =   4
      Top             =   6000
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6405
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppUpgradeSel.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16325
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "10:54"
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
   Begin VB.Frame fraMain 
      Height          =   5985
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   11280
      Begin VSFlex8Ctl.VSFlexGrid vsReport 
         Height          =   2415
         Left            =   0
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   11220
         _cx             =   19791
         _cy             =   4260
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   0   'False
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
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   0
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   14737632
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAppUpgradeSel.frx":0E1C
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         AutoSizeMouse   =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsOptional 
         Height          =   2535
         Left            =   0
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   11220
         _cx             =   19791
         _cy             =   4471
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   0   'False
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
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   0
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   14737632
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAppUpgradeSel.frx":0EDC
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid vsHis 
         Height          =   2715
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   11220
         _cx             =   19791
         _cy             =   4789
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   0   'False
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
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   0
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   14737632
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAppUpgradeSel.frx":0F92
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         AutoSizeMouse   =   0   'False
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
         Width           =   11280
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "确定本次要升迁的历史数据空间的用户（历史数据空间的所有者）、密码及服务器"
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
         Width           =   8100
      End
   End
End
Attribute VB_Name = "frmAppUpgradeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'====================================================================
'==变量
'====================================================================
'升级选择类型
Public Enum AupgradeSelType
    AST_His = 0 '历史库选择
    AST_OptProc = 1 '可选过程选择
    AST_Report = 2 '导入报表选择
End Enum
'历史库选择列
Private Enum HisCols
    HC_ID = 0
    HC_系统 = 1
    HC_HisDB = 2
    HC_IsCur = 3
    HC_Server = 4
    HC_PWD = 5
    HC_CurVer = 6
    HC_AimVer = 7
    HC_WarnInfo = 8
    HC_Sel = 9
End Enum
'可选过程选择列
Private Enum ProcCols
    PC_ID = 0
    PC_系统 = 1
    PC_ProcExector = 2
    PC_ProcVer = 3
    PC_ProcInfo = 4
    PC_Sel = 5
End Enum
'导入报表选择列
Private Enum ReportCols
    RC_ID = 0
    RC_系统 = 1
    RC_RptNo = 2
    RC_RptName = 3
    RC_AllImp = 4
    RC_SourceImp = 5
End Enum
Private mastSelType As AupgradeSelType '升级选择类型
Private mrsSource As ADODB.Recordset '初始界面所需要数据源，并记录界面项目的选择状态
Private mblnExecBef As Boolean '是否提前执行
Private mrsSysFiles As ADODB.Recordset '升迁需要执行的脚本记录集
Private mblnOk As Boolean
'====================================================================
'==公共接口
'====================================================================
Public Function ShowMe(frmParent As Object, ByVal astSelType As AupgradeSelType, Optional ByRef rsSource As ADODB.Recordset, Optional ByRef rsSysFiles As ADODB.Recordset, Optional ByVal blnExecBef As Boolean) As Boolean
'功能：展示选择界面
'参数： frmParent=父窗体
'           astSelType=升级选择类型
'           rsSource=初始界面所需要数据源
'返回：rsSource=界面选择状态
'         ShowMe=是否退出，暂时未使用

    mastSelType = astSelType
    rsSource.Filter = ""
    Set mrsSource = rsSource
    mblnExecBef = blnExecBef
    Set mrsSysFiles = rsSysFiles
    Me.Show 1, frmParent
    Set rsSource = mrsSource
    Set rsSysFiles = mrsSysFiles
    ShowMe = mblnOk
End Function
'====================================================================
'==控件事件
'====================================================================

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNotSel_Click()
    Call SetSelBeach
End Sub

Private Sub cmdSelAll_Click()
    Call SetSelBeach(True)
End Sub

Private Sub Form_Load()
    Call ApplyOEM(stbThis)
    vsHis.Visible = mastSelType = AST_His: vsHis.Enabled = mastSelType = AST_His
    vsOptional.Visible = mastSelType = AST_OptProc: vsOptional.Enabled = mastSelType = AST_OptProc
    vsReport.Visible = mastSelType = AST_Report: vsReport.Enabled = mastSelType = AST_Report
    lblInfo.Caption = Decode(mastSelType, AST_His, "确定本次要升迁的历史数据空间的用户（历史数据空间的所有者）、密码及服务器。", _
                                                                AST_OptProc, "请仔细阅读每个过程的说明，并结合实际情况，依次确认下面的哪些过程需要在本次升迁过程中自动执行。", _
                                                                AST_Report, "设置要自动导入的报表，可根据情况设置要导入的报表，也可以取消导入稍后手工处理。")
    Me.Caption = Decode(mastSelType, AST_His, "历史库验证以及选择", AST_OptProc, "可选过程选择", AST_Report, "报表导入设置")
    Call LoadData
End Sub

Private Sub Form_Resize()
    vsHis.Top = fraTop.Top + fraTop.Height + 60: vsOptional.Top = vsHis.Top: vsReport.Top = vsHis.Top
    vsHis.Left = 30: vsOptional.Left = 30: vsReport.Left = 30
    vsHis.Width = Me.ScaleWidth - 60: vsOptional.Width = vsHis.Width: vsReport.Width = vsHis.Width
    vsHis.Height = fraMain.Height - vsHis.Top - 30: vsOptional.Height = vsHis.Height: vsReport.Height = vsHis.Height
End Sub

Private Sub vsHis_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Select Case Col
        Case HC_PWD
            vsHis.Cell(flexcpData, Row, Col) = IIf(InStr(1, vsHis.TextMatrix(Row, Col), "*") <> 0, vsHis.Cell(flexcpData, Row, Col), vsHis.TextMatrix(Row, Col))
            vsHis.TextMatrix(Row, Col) = String(Len(vsHis.TextMatrix(Row, Col)), "*")
        Case HC_Sel
            Call RecUpdate(mrsSource, "ID=" & Val(vsHis.TextMatrix(Row, HC_ID)), "升级", IIf(Val(vsHis.TextMatrix(Row, HC_Sel)) = 0, 0, 1))
        End Select
        Call RefreshColor(Row)
End Sub

Private Sub vsHis_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = HC_PWD Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
           If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
              If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
                  If InStr(1, Chr(KeyAscii), "_") = 0 Then
                      If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then
                      Else
                          KeyAscii = 0
                      End If
                      Exit Sub
                  End If
              End If
           End If
        End If
        vsHis.Cell(flexcpData, Row, Col) = vsHis.Cell(flexcpData, Row, Col) & Chr(KeyAscii)
    End If
End Sub

Private Sub vsHis_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '刘兴宏:加入
    '设置编辑的密码
    If Col = HC_PWD Then
        SendMessage vsHis.EditWindow, EM_SETPASSWORDCHAR, Asc("*"), 0
    End If
End Sub

Private Sub vsHis_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '升级时，如果是当前的值，则该值不能更改
    Cancel = Col <> HC_Sel And Col <> HC_PWD And Col <> HC_Server Or Col = HC_Sel And Trim(vsHis.TextMatrix(Row, HC_IsCur)) <> ""
End Sub

Private Sub vsHis_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strUserName As String, strBakName As String, strPassword As String, strServer As String
    Dim cnTmp As ADODB.Connection, rsTmp As ADODB.Recordset
    Dim strFilter As String, strMaxVer As String
    Dim strDbLink As String
    
    Select Case Col
        Case HC_Server, HC_PWD
            '看是否输入是否有效
            If Col = HC_Server Then
                strPassword = vsHis.Cell(flexcpData, Row, HC_PWD)
                strServer = vsHis.EditText
            Else
                If InStr(1, vsHis.EditText, "*") > 0 Then
                    strPassword = vsHis.Cell(flexcpData, Row, Col)
                Else
                    strPassword = vsHis.EditText
                End If
                strServer = vsHis.TextMatrix(Row, HC_Server)
            End If
            strPassword = Trim(strPassword)
            strServer = UCase(Trim(strServer))
            strBakName = UCase(Trim(vsHis.TextMatrix(Row, HC_HisDB)))
            strUserName = UCase(Trim(vsHis.Cell(flexcpData, Row, HC_HisDB)))
            strDbLink = UCase(Trim(vsHis.Cell(flexcpData, Row, HC_Server)))
            '判断服务器是否改变，改变，则去掉验证标志以及相关脚本
            mrsSource.Filter = "ID=" & Val(vsHis.TextMatrix(Row, HC_ID))
            If strPassword = "" And mrsSource!密码 & "" <> "" Then '没有输入密码，取上次验证的密码
                strPassword = mrsSource!密码 & ""
            End If
            If strServer <> mrsSource!服务器 & "" Then
                If strServer <> "" Then '输入服务器了就重新验证
                    mrsSource.Update Array("升级", "服务器", "当前版本", "密码", "验证", "可升级", "中止信息", "检查结果", "可提前升级", "提前中止信息", "提前检查结果"), _
                                                 Array(mrsSource!升级 * -1, strServer, Null, Null, 0, 1, Null, Null, 1, Null, Null)
                    '删除脚本
                    Call RecDelete(mrsSysFiles, "系统编号=" & mrsSource!系统编号 & " And 所有者='" & strBakName & "'")
                Else '没有输入服务器就使用以前的服务器
                    strServer = mrsSource!服务器 & ""
                End If
            End If
            
            If strPassword <> "" And strUserName <> "" And strServer <> "" Then
                Set cnTmp = gobjRegister.GetConnection(strServer, strUserName, strPassword, False, OraOLEDB, "", False)
                If cnTmp.State = adStateOpen Then
                    Set rsTmp = ReadHisUpgrade(cnTmp, strUserName, True, , strDbLink <> "")
                    Call RecUpdate(mrsSource, "所有者='" & strUserName & "' And 服务器='" & strServer & "' And 验证=0", "验证", 1)
                    rsTmp.Sort = ""
                    If rsTmp.EOF Then
                        Call RecUpdate(mrsSource, "所有者='" & strUserName & "' And 服务器='" & strServer & "'", "密码", strPassword, "可升级", 0, "可提前升级", 0, "检查结果", "历史表空间数据结构缺失导致无法升级！")
                    Else
                        Do While Not rsTmp.EOF
                            mrsSource.Filter = "系统编号=" & rsTmp!系统编号 & " And 所有者='" & strUserName & "' And 服务器='" & strServer & "'"
                            Do While Not mrsSource.EOF
                                If mrsSource!验证 = 1 Then mrsSource.Update "验证", 2
                                mrsSource.Update Array("密码", "当前版本", "中止信息", "提前中止信息"), Array(strPassword, rsTmp!当前版本, rsTmp!中止信息, rsTmp!提前中止信息)
                                '判断能否升迁
                                If Not IsVerSion(rsTmp!当前版本 & "") Then
                                    mrsSource.Update Array("可升级", "检查结果", "可提前升级"), Array(0, "历史数据空间的版本不可识别。请检查！", 0)
                                ElseIf VerFull(rsTmp!当前版本 & "") >= VerFull(mrsSource!目标版本 & "") Then
                                    mrsSource.Update Array("可升级", "检查结果", "可提前升级"), Array(0, "历史数据空间的版本高于本次升迁目标版本，不能升迁！", 0)
                                Else
                                    Set mrsSysFiles = GetUpgradeFiles(mrsSysFiles, rsTmp!系统编号, rsTmp!当前版本, mrsSource!配置文件, rsTmp!中止信息, rsTmp!提前中止信息, mrsSource!目标版本, , strBakName)
                                    '获取提前执行的目标版本
                                    If mblnExecBef Then
                                        strFilter = "所有者='" & strBakName & "' And FileType=" & FT_Before
                                        mrsSysFiles.Filter = strFilter: mrsSysFiles.Sort = "FullSPVer Desc": strMaxVer = ""
                                        If Not mrsSysFiles.EOF Then
                                            strMaxVer = mrsSysFiles!SPVer
                                            mrsSysFiles.Filter = strFilter & " And 配置版本>'" & VerFull(rsTmp!当前版本 & "") & "'": mrsSysFiles.Sort = "FullSPVer"
                                            If Not mrsSysFiles.EOF Then
                                                mrsSysFiles.Filter = strFilter & " And FullSPVer<'" & mrsSysFiles!FullSPVer & "'": mrsSysFiles.Sort = "FullSPVer Desc"
                                                If Not mrsSysFiles.EOF Then
                                                    strMaxVer = mrsSysFiles!SPVer
                                                Else
                                                    strMaxVer = ""
                                                    mrsSource.Update Array("可提前升级", "提前检查结果"), Array(0, "没有可执行的提前升级脚本，不能提前升迁！")
                                                End If
                                            End If
                                        Else
                                            mrsSource.Update Array("可提前升级", "提前检查结果"), Array(0, "没有提前升级脚本，不能提前升迁！")
                                        End If
                                        mrsSource.Update "提前目标版本", strMaxVer
                                        '删除非提前执行脚本
                                        Call RecDelete(mrsSysFiles, "所有者='" & strBakName & "' And FileType<>" & FT_Before)
                                        '删除大于提前目标版本的提前升级脚本
                                        Call RecDelete(mrsSysFiles, strFilter & " And FullSPVer>'" & VerFull(strMaxVer) & "'")
                                    End If
                                End If
                                mrsSource.MoveNext
                            Loop
                            rsTmp.MoveNext
                        Loop
                    End If
                    '标记未在历史空间中注册
                    Call RecUpdate(mrsSource, "验证=1", "可升级", 0, "可提前升级", 0, "检查结果", "该系统的历史空间未在ZLBakInfo中注册！")
                Else
                    Cancel = True
                    Exit Sub
                End If
            End If
        Case HC_Sel
            Call RecUpdate(mrsSource, "ID=" & Val(vsHis.TextMatrix(Row, HC_ID)), "升级", IIf(Val(vsHis.TextMatrix(Row, HC_Sel)) = 0, 0, 1))
    End Select
    Call LoadData '重新加载数据
    Call RefreshColor(Row)
End Sub

Private Sub vsOptional_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsOptional
        If Col = PC_Sel Then
            Call RecUpdate(mrsSource, "ID=" & Val(vsOptional.TextMatrix(Row, PC_ID)), "执行", IIf(Val(vsOptional.TextMatrix(Row, PC_Sel)) = 0, 0, 1))
            Call RefreshColor(Row)
        End If
    End With
End Sub

Private Sub vsOptional_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> PC_Sel Then Cancel = True
End Sub

Private Sub vsReport_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsReport
        If Col = RC_AllImp Then
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                .TextMatrix(Row, RC_SourceImp) = 0  '不导入数据源
            End If
        ElseIf Col = RC_SourceImp Then
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                .TextMatrix(Row, RC_AllImp) = 0  '不整体导入
            End If
        End If
        Call RecUpdate(mrsSource, "ID=" & Val(.TextMatrix(Row, RC_ID)), "覆盖类型", IIf(Val(.TextMatrix(Row, RC_SourceImp)) <> 0, 2, IIf(Val(.TextMatrix(Row, RC_AllImp)), 1, 0)))
        Call RefreshColor(Row)
    End With
End Sub

Private Sub vsReport_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If vsReport.ColWidth(Col) < 500 Then vsReport.ColWidth(Col) = 500
End Sub

Private Sub vsReport_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = RC_AllImp Or Col = RC_SourceImp Then
        Cancel = True
    End If
End Sub

Private Sub vsReport_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = RC_AllImp Or Col = RC_SourceImp) Then
        Cancel = True
    End If
End Sub

'====================================================================
'==方法
'====================================================================
Private Sub SetSelBeach(Optional ByVal blnSel As Boolean)
'功能：设置批量选择
'参数：blnSel=True：批量选择，False:批量取消
    Dim intSel As Integer, lngCol As Long, lngOtherCol As Long
    Dim i As Long
    Dim vsTmp As VSFlexGrid
    
    intSel = IIf(blnSel, 1, 0): lngCol = -1
    If mastSelType = AST_Report Then
        If vsReport.Col = RC_AllImp Or vsReport.Col = RC_SourceImp Then
            lngCol = vsReport.Col
            lngOtherCol = IIf(lngCol = RC_AllImp, RC_SourceImp, RC_AllImp)
        End If
        Set vsTmp = vsReport
        Call RecUpdate(mrsSource, "", "覆盖类型", IIf(blnSel And lngCol = RC_SourceImp, 2, IIf(blnSel And lngCol = RC_AllImp, 1, 0)))
    ElseIf mastSelType = AST_His Then
        Set vsTmp = vsHis
        lngCol = HC_Sel
        Call RecUpdate(mrsSource, "可升级 = 1 And 可提前升级 = 1" & IIf(Not blnSel, " And 当前<>1", ""), "升级", IIf(blnSel, 1, 0))
    Else
        Set vsTmp = vsOptional
        lngCol = PC_Sel
        Call RecUpdate(mrsSource, "", "执行", IIf(blnSel, 1, 0))
    End If
    If lngCol = -1 Then Exit Sub
    With vsTmp
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, 0) = 1 Then
                .TextMatrix(i, lngCol) = intSel
                '报表导入只能选择一种导入方法，选择了一种，就取消另一种
                If intSel = 1 And mastSelType = AST_Report Then
                    .TextMatrix(i, lngOtherCol) = 0
                End If
            End If
        Next
    End With
    Call RefreshColor
End Sub

Private Sub LoadData()
    Dim vsTmp As VSFlexGrid
    Dim strPre As String, strPrePOther As String
    
    mrsSource.Filter = ""
    If mastSelType = AST_His Then
        Set vsTmp = vsHis
        mrsSource.Sort = "系统编号,当前,编号,ID"
    ElseIf mastSelType = AST_OptProc Then
        Set vsTmp = vsOptional
        mrsSource.Sort = "系统编号, 历史库, 执行者,ID"
    Else
        Set vsTmp = vsReport
        mrsSource.Sort = "系统编号,编号,ID"
    End If
    With vsTmp
        .Rows = .FixedRows
        .MergeCompare = flexMCTrimNoCase
        .MergeCells = flexMergeRestrictColumns
        Select Case mastSelType
            Case AST_His
                Do While Not mrsSource.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, HC_ID) = mrsSource!Id
                    .Cell(flexcpData, .Rows - 1, HC_ID) = IIf(Not mblnExecBef And mrsSource!可升级 = 1 Or mblnExecBef And mrsSource!可提前升级 = 1, 1, 0)
                    If .Cell(flexcpData, .Rows - 1, HC_ID) = 1 And Val(mrsSource!当前 & "") = 1 Then .Cell(flexcpData, .Rows - 1, HC_ID) = -1
                    .TextMatrix(.Rows - 1, HC_系统) = mrsSource!系统名称 & ""
                    .TextMatrix(.Rows - 1, HC_HisDB) = mrsSource!名称 & ""
                    .Cell(flexcpData, .Rows - 1, HC_HisDB) = mrsSource!所有者 & ""
                    .TextMatrix(.Rows - 1, HC_IsCur) = IIf(Val(mrsSource!当前 & "") = 1, "√", "")
                    .TextMatrix(.Rows - 1, HC_CurVer) = mrsSource!当前版本 & ""
                    .TextMatrix(.Rows - 1, HC_AimVer) = IIf(mblnExecBef, mrsSource!提前目标版本 & "", mrsSource!目标版本 & "")
                    .Cell(flexcpData, .Rows - 1, HC_PWD) = mrsSource!密码 & ""
                    .TextMatrix(.Rows - 1, HC_PWD) = String(Len(mrsSource!密码 & ""), "*")
                    .TextMatrix(.Rows - 1, HC_Sel) = Val(mrsSource!升级 & "")
                    .TextMatrix(.Rows - 1, HC_Server) = mrsSource!服务器 & ""
                    .Cell(flexcpData, .Rows - 1, HC_Server) = mrsSource!DB连接 & ""
                    .TextMatrix(.Rows - 1, HC_WarnInfo) = IIf(mrsSource!检查结果 & "" = "", mrsSource!提前检查结果 & "", mrsSource!检查结果)
                    .RowData(.Rows - 1) = 0
                    mrsSource.MoveNext
                Loop
            Case AST_OptProc
                Do While Not mrsSource.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, PC_ID) = mrsSource!Id
                    .Cell(flexcpData, .Rows - 1, PC_ID) = 1
                    .TextMatrix(.Rows - 1, PC_系统) = mrsSource!系统名称 & ""
                    .TextMatrix(.Rows - 1, PC_ProcExector) = mrsSource!执行者 & ""
                    .TextMatrix(.Rows - 1, PC_ProcInfo) = mrsSource!名称 & vbNewLine & mrsSource!注释
                    .TextMatrix(.Rows - 1, PC_ProcVer) = mrsSource!SPVer & ""
                    .TextMatrix(.Rows - 1, PC_Sel) = Val(mrsSource!执行 & "")
                     .RowData(.Rows - 1) = Val(mrsSource!执行 & "")
                    mrsSource.MoveNext
                Loop
                Call vsTmp.AutoSize(PC_ProcInfo)
            Case Else
                Do While Not mrsSource.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, RC_ID) = mrsSource!Id
                    .Cell(flexcpData, .Rows - 1, RC_ID) = 1
                    .TextMatrix(.Rows - 1, RC_系统) = mrsSource!系统名称 & ""
                    .TextMatrix(.Rows - 1, RC_RptNo) = mrsSource!编号 & ""
                    .TextMatrix(.Rows - 1, RC_RptName) = mrsSource!名称 & ""
                    .TextMatrix(.Rows - 1, RC_AllImp) = IIf(Val(mrsSource!覆盖类型 & "") = 1, 1, 0)
                    .TextMatrix(.Rows - 1, RC_SourceImp) = IIf(Val(mrsSource!覆盖类型 & "") = 2, 1, 0)
                     .RowData(.Rows - 1) = Val(mrsSource!覆盖类型 & "")
                    mrsSource.MoveNext
                Loop
        End Select
        
        If mastSelType = AST_His Then
            .MergeCol(HC_系统) = True
            .MergeCol(HC_HisDB) = True
        ElseIf mastSelType = AST_OptProc Then
            .MergeCol(PC_系统) = True
            .MergeCol(PC_ProcExector) = True
        Else
            .MergeCol(RC_系统) = True
        End If
    End With
    Call RefreshColor
End Sub

Private Sub RefreshColor(Optional ByVal lngRow As Long)
    Dim i As Long
    
    If mastSelType = AST_His Then
        With vsHis
            If lngRow < .FixedRows Then
                For i = .FixedRows To .Rows - 1
                    If Val(.Cell(flexcpData, i, HC_ID)) = 0 Then
                        .Cell(flexcpForeColor, i, HC_系统, i, .Cols - 1) = &H2222B2 '火砖红
                    Else
                        .Cell(flexcpForeColor, i, HC_系统, i, .Cols - 1) = .ForeColor
                    End If
                Next
            Else
                If Val(.Cell(flexcpData, lngRow, HC_ID)) = 0 Then
                    .Cell(flexcpForeColor, lngRow, HC_系统, lngRow, .Cols - 1) = &H2222B2 '火砖红
                Else
                    .Cell(flexcpForeColor, lngRow, HC_系统, lngRow, .Cols - 1) = .ForeColor
                End If
            End If
        End With
    ElseIf mastSelType = AST_OptProc Then
    
    Else
        With vsReport
            If lngRow < .FixedRows Then
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, RC_AllImp)) <> 0 Then
                        .Cell(flexcpForeColor, i, HC_系统, i, .Cols - 1) = .ForeColor
                    ElseIf Val(.TextMatrix(i, RC_SourceImp)) <> 0 Then
                        .Cell(flexcpForeColor, i, HC_系统, i, .Cols - 1) = vbBlue
                    Else
                        .Cell(flexcpForeColor, i, HC_系统, i, .Cols - 1) = &H808080   '灰色
                    End If
                Next
            Else
                If Val(.TextMatrix(lngRow, RC_AllImp)) <> 0 Then
                    .Cell(flexcpForeColor, lngRow, HC_系统, lngRow, .Cols - 1) = .ForeColor
                ElseIf Val(.TextMatrix(lngRow, RC_SourceImp)) <> 0 Then
                    .Cell(flexcpForeColor, lngRow, HC_系统, lngRow, .Cols - 1) = vbBlue
                Else
                    .Cell(flexcpForeColor, lngRow, HC_系统, lngRow, .Cols - 1) = &H808080  '灰色
                End If
            End If
        End With
    End If
End Sub

