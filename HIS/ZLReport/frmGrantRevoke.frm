VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGrantRevoke 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11295
   Icon            =   "frmGrantRevoke.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList img16 
      Left            =   2880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrantRevoke.frx":058A
            Key             =   "SYS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrantRevoke.frx":0B24
            Key             =   "MODULE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrantRevoke.frx":10BE
            Key             =   "REPORT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrantRevoke.frx":1458
            Key             =   "NODE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrantRevoke.frx":17F2
            Key             =   "NEW"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwReports 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1508
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSingleClear 
      Caption         =   "<"
      Height          =   360
      Left            =   4080
      TabIndex        =   9
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton cmdSingleSelect 
      Caption         =   ">"
      Height          =   360
      Left            =   4080
      TabIndex        =   8
      Top             =   3120
      Width           =   255
   End
   Begin VB.Frame fraSelectted 
      Caption         =   "发布位置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   4440
      TabIndex        =   10
      Top             =   1320
      Width           =   6735
      Begin VB.CommandButton cmdClear 
         Caption         =   "清空(&L)"
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Top             =   4395
         Width           =   1110
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存(&S)"
         Height          =   360
         Left            =   4320
         TabIndex        =   13
         Top             =   4395
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   360
         Left            =   5520
         TabIndex        =   14
         Top             =   4395
         Width           =   1110
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSelected 
         Height          =   3975
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   6495
         _cx             =   11456
         _cy             =   7011
         Appearance      =   0
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
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
      End
   End
   Begin VB.Frame fraSelecting 
      Caption         =   "可选位置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   840
         TabIndex        =   7
         ToolTipText     =   "Enter键：查找；F3键：继续查找"
         Top             =   4440
         Width           =   2895
      End
      Begin MSComctlLib.TreeView tvwSelecting 
         Height          =   3615
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   6376
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.ComboBox cboMenuGroup 
         Appearance      =   0  'Flat
         Height          =   300
         ItemData        =   "frmGrantRevoke.frx":1B8C
         Left            =   960
         List            =   "frmGrantRevoke.frx":1B8E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   2775
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查找(&F)"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   4470
         Width           =   630
      End
      Begin VB.Label lblMenuGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "菜单组别"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Label lblReportName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "报表名"
      Height          =   180
      Left            =   1320
      TabIndex        =   16
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   180
      Left            =   4200
      TabIndex        =   15
      Top             =   6200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblReports 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "待发布报表："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1170
   End
End
Attribute VB_Name = "frmGrantRevoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum enuMode
    导航台 = 0
    模块
End Enum

Private Const MSTR_NAVIGATION As String = _
    "新,,3,300|组别,,3,800|系统,,3,1500|菜单,,3,1500|报表,,3,2500|PID,,0,0,n|MenuID,,0,0,n|ReportID,,0,0,n|" & _
    "程序ID,,0,0,n|SysNo,,0,0,n"

Private WithEvents mobjSelected As clsVSFlexGridEx
Attribute mobjSelected.VB_VarHelpID = -1
Private mbytMode As Byte
Private mblnResult As Boolean
Private mblnGroup As Boolean
Private mcolRevoke As Collection
Private mlngPrevious As Long
Private mblnChanged As Boolean

Property Get Mode_() As enuMode
    Mode_ = mbytMode
End Property
Property Let Mode_(ByVal bytMode As enuMode)
    mbytMode = bytMode
End Property

Public Function ShowMe(ByVal frmOwner As Form, ByVal vsfSelect As VSFlexGrid) As Boolean
    Dim lngCount As Long
    
    '初始化
    mblnGroup = UCase$(vsfSelect.name) = "VSFGROUP"
    If Mode_ = 模块 And mblnGroup Then
        MsgBox "组报表不允许模块类的发布管理！", vbInformation, App.Title
        Exit Function
    End If
    
    Call InitReportList(vsfSelect, lngCount)
    If lngCount <= 0 Then
        MsgBox "未选择报表！", vbInformation, App.Title
        Exit Function
    End If
    Call InitMenuGroup
    Call InitSelecting
    Call InitSelected
    
    If Mode_ = 导航台 Then
        Me.Caption = "发布管理-导航台"
    Else
        Me.Caption = "发布管理-模块"
    End If
    
    '加载数据
    Call RefreshSelected
    Call RefreshMenuGroup
    
    '窗体
    Me.Show vbModal, frmOwner
    ShowMe = mblnResult
    If mblnResult Then
        Unload Me
    End If
End Function

Private Sub cboMenuGroup_Click()
    If Me.Visible = False Then Exit Sub
    
    Call RefreshMenuGroup
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim lngRow As Long
    
    If mblnChanged = False Then mblnChanged = True
    
    With vsfSelected
        .Redraw = False
        For lngRow = .Rows - 1 To 1 Step -1
            '记录已发布的菜单结点信息
            If .Cell(flexcpData, lngRow, .ColIndex("新")) = Val("0-已发布") Then
                Call RevokeAdd(lngRow)
            End If
            .RemoveItem lngRow
        Next
        .Redraw = True
    End With
    
    With tvwSelecting
        For lngRow = 1 To .Nodes.count
            .Nodes(lngRow).Bold = False
        Next
    End With
End Sub

Private Sub DBObjectPrivs(ByVal strFunc As String, ByVal lngReportID As Long, ByVal lngProgID As Long _
    , ByVal lngSys As Long, ByRef colSQL As Collection)
    
    Dim strObject As String, strOwner As String, strSQL As String, strTmp As String
    Dim i As Long
    Dim arrTmp() As String
    
    strObject = GetReportObjects(lngReportID)
    If strObject <> "" Then
        strObject = Mid$(strObject, 2)
        arrTmp = Split(strObject, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            strOwner = Left$(arrTmp(i), InStr(arrTmp(i), ".") - 1)
            If InStr(";SYS;SYSTEM;ZLTOOLS;", ";" & strOwner & ";") <= 0 Then
                strTmp = Mid$(arrTmp(i), InStr(arrTmp(i), ".") + 1)
                strSQL = GetInsertProgPrivs(lngSys, lngProgID, strFunc, strTmp, strOwner, "SELECT")
                Call AddArray(colSQL, strSQL)
            End If
        Next
    End If
End Sub

Private Function ExistGrantData(ByVal blnGroup As Boolean, ByVal lngID As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    If Mode_ = 导航台 Then
        If blnGroup Then
            strSQL = _
                "Select Count(1) Rec " & vbCr & _
                "From zlMenus A, zlRPTGroups B " & vbCr & _
                "Where a.模块 = b.程序id And a.系统 Is Null And b.ID = [1]"
        Else
            strSQL = _
                "Select Count(1) Rec " & vbCr & _
                "From zlMenus A, zlReports B " & vbCr & _
                "Where a.模块 = b.程序id And a.系统 Is Null And b.ID = [1]"
        End If
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取指定报表发布至导航台的记录", lngID)
    Else
        strSQL = "Select Count(1) Rec From zlRPTPuts Where 报表ID = [1]"
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取指定报表发布至模块的记录", lngID)
    End If
    ExistGrantData = rsTemp!Rec > 1
    rsTemp.Close
    
    Exit Function
    
hErr:
    If mdlPublic.ErrCenter = 1 Then Resume
End Function

Private Sub cmdSave_Click()
'说明：
'  1.独立报表、子报表发布到导航台。
'    a.首次发布需要处理“zlReports、zlMenus、zlPrograms、zlProgFuncs、zlProgPrivs”表数据；
'    b.其次发布只需要处理“zlReports、zlMenus”表数据；
'  2.组报表发布到导航台。
'    a.首次发布需要处理“zlRPTGroups、zlRPTSubs、zlMenus、zlPrograms、zlProgFuncs、zlProgPrivs”表数据；
'    b.其次发布只需要处理“zlReports、zlMenus”表数据；
'  3.独立报表、子报表发布到模块。
'    发布需要处理“zlReport、zlRPTPuts、zlProgFuncs、zlProgPrivs”表数据；
'  4.
'注意：
'  1.系统报表不允许发布操作，只允许自定义报表发布操作；
'  2.组报表不允许发布到模块；

    Dim colSQL As Collection, colReportGroup As Collection
    Dim l As Long, k As Long, lngSys As Long
    Dim lngGroupID As Long, lngMenuID As Long, lngReportID As Long, lngProgID As Long, lngPid As Long
    Dim strVerifiedRID As String, strSQL As String, strTmp As String, strMenuGroup As String
    Dim rsTmp As ADODB.Recordset
    Dim arrID() As String
    Dim blnTrans As Boolean
    Dim ppbItem As PropertyBag

    With vsfSelected
        '发布位置过多提醒，但不禁止
        For l = 1 To .Rows - 1
            If .Cell(flexcpData, l, .ColIndex("新")) = Val("1-新增") Then
                k = k + 1
            End If
        Next
        If k > 5 Then
            If MsgBox("提醒：" & vbCr & "发布的项目过多，确认继续发布吗？", vbQuestion + vbYesNo + vbDefaultButton2 _
                , App.Title) = vbNo Then
                Exit Sub
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        
        On Error GoTo hErr
    
        '检查报表（独立、组、子）权限和密码
        Set colReportGroup = New Collection
        For l = 1 To .Rows - 1
            If .Cell(flexcpData, l, .ColIndex("新")) = Val("1-新增") Then
                If mblnGroup Then
                    '组报表
                    lngGroupID = Val(.TextMatrix(l, .ColIndex("ReportID")))
                    If InStr(strVerifiedRID & ",", "," & lngGroupID & ",") <= 0 Then
                        strTmp = ""
                        strSQL = _
                            "Select Distinct a.Id, a.名称 " & vbCr & _
                            "From zlReports A, zlRPTSubs B " & vbCr & _
                            "Where a.Id = b.报表id And b.组id = [1] "
                        Set rsTmp = mdlPublic.OpenSQLRecord(strSQL, "获取报表组的报表ID", lngGroupID)
                        Do While rsTmp.EOF = False
                            If ReportVerify(rsTmp!id, rsTmp!名称) = False Then
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            strTmp = strTmp & "," & rsTmp!id & ";" & rsTmp!名称
                            rsTmp.MoveNext
                        Loop
                        '记录组报表ID
                        strVerifiedRID = strVerifiedRID & "," & lngGroupID
                        '记录组报表的子报表ID和子报表名称
                        If strTmp <> "" Then
                            colReportGroup.Add Mid$(strTmp, 2), "_" & lngGroupID
                        End If
                        rsTmp.Close
                    End If
                Else
                    '独立报表或子报表
                    lngReportID = Val(.TextMatrix(l, .ColIndex("ReportID")))
                    If InStr(strVerifiedRID & ",", "," & lngReportID & ",") <= 0 Then
                        If ReportVerify(lngReportID, .TextMatrix(l, .ColIndex("报表"))) = False Then
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        '记录报表ID
                        strVerifiedRID = strVerifiedRID & "," & lngReportID
                    End If
                End If
            End If
        Next
        
        If Mode_ = 导航台 Then
            '1.导航台
                        
            If mblnGroup Then
                '组报表
                
                '发布
                For l = 1 To .Rows - 1
                    If .Cell(flexcpData, l, .ColIndex("新")) = Val("1-新增") Then
                        Set colSQL = New Collection
                        strMenuGroup = Trim$(.TextMatrix(l, .ColIndex("组别")))
                        lngPid = Val(.TextMatrix(l, .ColIndex("PID")))
                        lngMenuID = Val(.TextMatrix(l, .ColIndex("MenuID")))
                        lngGroupID = Val(.TextMatrix(l, .ColIndex("ReportID")))
                        lngProgID = Val(.TextMatrix(l, .ColIndex("程序ID")))
                        If lngProgID <= 0 Then
                            lngProgID = GetProgID(lngGroupID, True)     '报表的程序ID是否已存在（实时）
                            If lngProgID <= 0 Then
                                lngProgID = mdlPublic.GetNewProgID()    '生成新的程序ID
                            
                                strSQL = _
                                    "Update zlRPTSubs A Set 功能 = (Select 名称 From zlReports Where ID = A.报表ID) " & vbCr & _
                                    "Where 组ID = " & lngGroupID
                                Call AddArray(colSQL, strSQL)
                                
                                strSQL = _
                                    "Update zlRPTGroups " & vbCr & _
                                    "Set 程序ID = " & lngProgID & ", 发布时间 = Sysdate " & vbCr & _
                                    "Where ID = " & lngGroupID
                                Call AddArray(colSQL, strSQL)
                                
                                strSQL = _
                                    "Insert Into zlPrograms(序号,标题,说明,系统,部件) " & vbCr & _
                                    "Select " & lngProgID & vbCr & _
                                    "" & _
                                    ", 名称, 说明" & _
                                    ", " & IIF(glngSys <= 0, "Null", glngSys) & _
                                    ", 'zl9Report' " & vbCr & _
                                    "From zlRPTGroups Where ID = " & lngGroupID
                                Call AddArray(colSQL, strSQL)
                                
                                strSQL = _
                                    "Insert Into zlProgFuncs(系统,序号,功能,说明)" & vbCr & _
                                    "Select " & IIF(glngSys <= 0, "Null", glngSys) & _
                                    ", " & lngProgID & _
                                    ", 名称, 说明 " & vbCr & _
                                    "From zlReports " & vbCr & _
                                    "Where ID In (Select 报表ID From zlRPTSubs Where 组ID = " & lngGroupID & ")"
                                Call AddArray(colSQL, strSQL)
                                
                                '组报表中所有子报表数据源的数据表对象访问权限
                                If CollectionFind(colReportGroup, "_" & lngGroupID) Then
                                    arrID = Split(colReportGroup("_" & lngGroupID), ",")
                                    For k = LBound(arrID) To UBound(arrID)
                                        strTmp = arrID(k)
                                        lngReportID = Val(strTmp)
                                        strTmp = Mid$(strTmp, InStr(strTmp, ";") + 1)
                                        Call DBObjectPrivs(strTmp, lngReportID, lngProgID, glngSys, colSQL)
                                    Next
                                End If
                            Else
                                strSQL = "Update zlRPTGroups Set 发布时间 = Sysdate Where ID = " & lngGroupID
                                Call AddArray(colSQL, strSQL)
                            End If
                        Else
                            strSQL = "Update zlRPTGroups Set 发布时间 = Sysdate Where ID = " & lngGroupID
                            Call AddArray(colSQL, strSQL)
                        End If
                        
                        strSQL = _
                            "Insert Into zlMenus(组别,ID,上级ID,标题,快键,说明,系统,模块,短标题,图标) " & vbCr & _
                            "Select '" & strMenuGroup & "'" & _
                            ", zlMenus_ID.Nextval" & _
                            ", " & lngPid & _
                            ", 名称, Null, 说明" & _
                            ", " & IIF(glngSys <= 0, "Null", glngSys) & _
                            ", " & lngProgID & _
                            ", 名称, 105 " & vbCr & _
                            "From zlRPTGroups Where ID = " & lngGroupID
                        Call AddArray(colSQL, strSQL)
                        
                        '执行发布的DML
                        gcnOracle.BeginTrans: blnTrans = True
                        For k = 1 To colSQL.count
                            'Debug.Print colSQL(k) & ";"
                            gcnOracle.Execute colSQL(k)
                        Next
                        gcnOracle.CommitTrans: blnTrans = False
                    End If
                Next
            Else
                '独立报表、子报表
                
                '发布
                For l = 1 To .Rows - 1
                    If .Cell(flexcpData, l, .ColIndex("新")) = Val("1-新") Then
                        Set colSQL = New Collection
                        strMenuGroup = Trim$(.TextMatrix(l, .ColIndex("组别")))
                        lngPid = Val(.TextMatrix(l, .ColIndex("PID")))
                        lngMenuID = Val(.TextMatrix(l, .ColIndex("MenuID")))
                        lngReportID = Val(.TextMatrix(l, .ColIndex("ReportID")))
                        lngProgID = Val(.TextMatrix(l, .ColIndex("程序ID")))
                        If lngProgID <= 0 Then
                            lngProgID = GetProgID(lngReportID)          '报表的程序ID是否已存在（实时）
                            If lngProgID <= 0 Then
                                lngProgID = mdlPublic.GetNewProgID()    '生成新的程序ID
                            
                                strSQL = _
                                    "Update zlReports " & vbCr & _
                                    "Set 功能 = '基本', 程序ID = " & lngProgID & ", 发布时间 = Sysdate " & vbCr & _
                                    "Where ID = " & lngReportID
                                Call AddArray(colSQL, strSQL)
                                
                                strSQL = _
                                    "Insert Into zlPrograms(序号,标题,说明,系统,部件) " & vbCr & _
                                    "Select " & lngProgID & _
                                    ", 名称, 说明" & _
                                    ", " & IIF(glngSys <= 0, "Null", glngSys) & _
                                    ", 'zl9Report' " & vbCr & _
                                    "From zlReports Where ID = " & lngReportID
                                Call AddArray(colSQL, strSQL)
                                
                                strSQL = _
                                    "Insert Into zlProgFuncs(系统,序号,功能) " & vbCr & _
                                    "Values (" & IIF(glngSys <= 0, "Null", glngSys) & _
                                    ", " & lngProgID & _
                                    ", '基本')"
                                Call AddArray(colSQL, strSQL)
                                
                                '组报表中所有子报表数据源的数据表对象访问权限
                                Call DBObjectPrivs("基本", lngReportID, lngProgID, glngSys, colSQL)
                            Else
                                strSQL = "Update zlReports Set 发布时间 = Sysdate Where ID = " & lngReportID
                                Call AddArray(colSQL, strSQL)
                            End If
                        Else
                            strSQL = "Update zlReports Set 发布时间 = Sysdate Where ID = " & lngReportID
                            Call AddArray(colSQL, strSQL)
                        End If
                        
                        strSQL = _
                            "Insert Into zlMenus(组别,ID,上级ID,标题,快键,说明,系统,模块,短标题,图标) " & vbCr & _
                            "Select '" & strMenuGroup & "'" & _
                            ", zlMenus_ID.Nextval" & _
                            ", " & lngPid & _
                            ", 名称, Null, 说明" & _
                            ", " & IIF(glngSys <= 0, "Null", glngSys) & _
                            ", " & lngProgID & _
                            ", 名称, 105 " & vbCr & _
                            "From zlReports Where ID = " & lngReportID
                        Call AddArray(colSQL, strSQL)
                        
                        '执行发布的DML
                        gcnOracle.BeginTrans: blnTrans = True
                        For k = 1 To colSQL.count
                            'Debug.Print colSQL(k) & ";"
                            gcnOracle.Execute colSQL(k)
                        Next
                        gcnOracle.CommitTrans: blnTrans = False
                    End If
                Next
            End If
            
            '撤销发布
            Set colSQL = New Collection
            If Not mcolRevoke Is Nothing Then
                For l = 1 To mcolRevoke.count
                    Set ppbItem = mcolRevoke(l)
                    lngProgID = ppbItem.ReadProperty("ProgID")
                    lngMenuID = ppbItem.ReadProperty("MenuID")
                    lngReportID = ppbItem.ReadProperty("ReportID")
                    
                    '判断报表是否存在发布数据（实时）
                    If ExistGrantData(mblnGroup, lngReportID) Then
                        If mblnGroup Then
                            '组报表还存在发布至导航台的数据
                            strSQL = "Update zlRPTGroups Set 发布时间 = Sysdate Where ID = " & lngReportID & " And 发布时间 <> Sysdate "
                            Call AddArray(colSQL, strSQL)
                        Else
                            '报表还存在发布至导航台的数据
                            strSQL = _
                                "Update zlReports Set 发布时间 = Sysdate " & vbCr & _
                                "Where ID = " & lngReportID & " And 发布时间 <> Sysdate "
                            Call AddArray(colSQL, strSQL)
                        End If
                        strSQL = "Delete From zlMenus Where ID = " & lngMenuID & " And Nvl(系统, 0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                    Else
                        If mblnGroup Then
                            '组报表未存在发布至导航台的数据
                            strSQL = _
                                "Update zlRPTGroups Set 程序ID = Null, 发布时间 = Null, 是否停用 = Null " & vbCr & _
                                "Where ID = " & lngReportID
                            Call AddArray(colSQL, strSQL)
                            
                            strSQL = _
                                "Update zlRPTSubs Set 功能 = Null " & vbCr & _
                                "Where 组ID = " & lngReportID
                            Call AddArray(colSQL, strSQL)
                        Else
                            '报表未存在发布至导航台的数据
                            strSQL = _
                                "Update zlReports " & vbCr & _
                                "Set 功能 = Null, 程序ID = Null, 是否停用 = Null" & vbCr & _
                                "  , 发布时间 = Case When Exists(Select 1 From zlRPTPuts Where 报表Id = " & lngReportID & ") Then " & vbCr & _
                                "                      Sysdate" & vbCr & _
                                "              Else Null End" & vbCr & _
                                "Where ID = " & lngReportID
                            Call AddArray(colSQL, strSQL)
                        End If
                        strSQL = "Delete From zlMenus Where 模块 = " & lngProgID & " And Nvl(系统,0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                        
                        strSQL = "Delete From zlProgPrivs Where 序号 = " & lngProgID & " And Nvl(系统,0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                        
                        strSQL = "Delete From zlProgFuncs Where 序号 = " & lngProgID & " And Nvl(系统,0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                        
                        strSQL = "Delete From zlPrograms Where 序号 = " & lngProgID & " And Nvl(系统,0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                        
                        '同时回收该表的数据（管理工具的角色授权会生成数据）
                        strSQL = "Delete From zlRoleGrant Where 序号 = " & lngProgID & " And Nvl(系统,0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                    End If
                    
                    '执行撤销发布的DML
                    gcnOracle.BeginTrans: blnTrans = True
                    For k = 1 To colSQL.count
                        'Debug.Print colSQL(k) & ";"
                        gcnOracle.Execute colSQL(k)
                    Next
                    gcnOracle.CommitTrans: blnTrans = False
                Next
            End If
        Else
            '2.模块
            
            '发布
            Set colSQL = New Collection
            For l = 1 To .Rows - 1
                If .Cell(flexcpData, l, .ColIndex("新")) = Val("1-新") Then
                    lngReportID = Val(.TextMatrix(l, .ColIndex("ReportID")))
                    lngProgID = Val(.TextMatrix(l, .ColIndex("PID")))
                    lngSys = Val(.TextMatrix(l, .ColIndex("SysNo")))
                
                    strSQL = "Update zlReports Set 发布时间 = Sysdate Where ID = " & lngReportID
                    Call AddArray(colSQL, strSQL)
                    
                    strSQL = _
                        "Insert Into zlRPTPuts(报表ID, 系统, 程序ID, 功能) " & vbCr & _
                        "Select " & lngReportID & _
                        ", " & IIF(lngSys <= 0, "Null", lngSys) & _
                        ", " & lngProgID & _
                        ", 名称 " & vbCr & _
                        "From zlReports Where ID = " & lngReportID
                    Call AddArray(colSQL, strSQL)
                
                    strSQL = _
                        "Insert Into zlProgFuncs(系统, 序号, 功能, 说明) " & vbCrLf & _
                        "Select " & IIF(lngSys <= 0, "Null", lngSys) & _
                        ", " & lngProgID & _
                        ", 名称, 说明 " & _
                        "From zlReports Where ID = " & lngReportID
                    Call AddArray(colSQL, strSQL)
                    
                    '独立报表、子报表数据源的数据表对象访问权限
                    Call DBObjectPrivs(Trim$(.TextMatrix(l, .ColIndex("报表"))), lngReportID, lngProgID, lngSys, colSQL)
                End If
            Next
            
            '执行发布的DML
            If colSQL.count > 0 Then
                gcnOracle.BeginTrans: blnTrans = True
                For k = 1 To colSQL.count
                    'Debug.Print colSQL(k) & ";"
                    gcnOracle.Execute colSQL(k)
                Next
                gcnOracle.CommitTrans: blnTrans = False
            End If
            
            '撤销发布
            Set colSQL = New Collection
            If Not mcolRevoke Is Nothing Then
                For l = 1 To mcolRevoke.count
                    Set ppbItem = mcolRevoke(l)
                    lngProgID = ppbItem.ReadProperty("ProgID")
                    lngReportID = ppbItem.ReadProperty("ReportID")
                    lngSys = ppbItem.ReadProperty("SysNo")
                    
                    strSQL = _
                        "Delete From zlRPTPuts " & vbCr & _
                        "Where 报表ID = " & lngReportID & " And 系统 = " & lngSys & " And 程序ID = " & lngProgID
                    Call AddArray(colSQL, strSQL)
                    
                    strSQL = _
                        "Delete From zlProgPrivs " & vbCr & _
                        "Where 系统 = " & lngSys & " And 序号 = " & lngProgID & vbCr & _
                        "  And 功能 = (Select 名称 From zlReports Where ID = " & lngReportID & ")"
                    Call AddArray(colSQL, strSQL)
                    
                    strSQL = _
                        "Delete From zlProgFuncs Where 系统 = " & lngSys & " And 序号=" & lngProgID & vbCr & _
                        "  And 功能 = (Select 名称 From zlReports Where ID = " & lngReportID & ")"
                    Call AddArray(colSQL, strSQL)
                    
                    strSQL = _
                        "Delete From zlRoleGrant Where 系统 = " & lngSys & " And 序号=" & lngProgID & vbCr & _
                        "  And 功能 = (Select 名称 From zlReports Where ID = " & lngReportID & ")"
                    Call AddArray(colSQL, strSQL)
                                        
                    '判断报表是否存在发布数据（实时）
                    If ExistGrantData(False, lngReportID) Then
                        '存在（暂无数据处理）
                    Else
                        '不存在
                        strSQL = _
                            "Update zlReports Set 发布时间 = NULL, 是否停用 = NULL " & vbCr & _
                            "Where 程序ID Is Null And ID = " & lngReportID
                        Call AddArray(colSQL, strSQL)
                    End If
                    
                    '执行撤销发布的DML
                    gcnOracle.BeginTrans: blnTrans = True
                    For k = 1 To colSQL.count
                        'Debug.Print colSQL(k) & ";"
                        gcnOracle.Execute colSQL(k)
                    Next
                    gcnOracle.CommitTrans: blnTrans = False
                Next
            End If
        End If
    End With
    
    mblnResult = True
    Me.Hide
    Screen.MousePointer = vbDefault
    Exit Sub

hErr:
    Screen.MousePointer = vbDefault
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    Call mdlPublic.ErrCenter
End Sub

Private Sub RevokeAdd(ByVal lngRow As Long)
    Dim strKey As String
    Dim ppbItem As PropertyBag
    
    Set ppbItem = New PropertyBag
    With vsfSelected
        strKey = Trim$(.TextMatrix(lngRow, .ColIndex("PID"))) & "_" & Trim$(.TextMatrix(lngRow, .ColIndex("ReportID")))
        Call ppbItem.WriteProperty("PID", Val(.TextMatrix(lngRow, .ColIndex("PID"))))
        Call ppbItem.WriteProperty("MenuID", Val(.TextMatrix(lngRow, .ColIndex("MenuID"))))
        Call ppbItem.WriteProperty("ReportID", Val(.TextMatrix(lngRow, .ColIndex("ReportID"))))
        Call ppbItem.WriteProperty("ProgID", Val(.TextMatrix(lngRow, .ColIndex("程序ID"))))
        Call ppbItem.WriteProperty("SysNo", Val(.TextMatrix(lngRow, .ColIndex("SysNo"))))
        Call CollectionAdd(mcolRevoke, strKey, ppbItem)
    End With
End Sub

Private Sub cmdSingleClear_Click()
    Dim lngRow As Long, lngID As Long
    
    If vsfSelected.Rows <= 1 Then Exit Sub
    If vsfSelected.SelectedRows <= 0 Then Exit Sub
    If vsfSelected.Row <= 0 Then Exit Sub
    
    If mblnChanged = False Then mblnChanged = True
    
    With vsfSelected
        lngRow = .Row
        lngID = Val(.TextMatrix(lngRow, .ColIndex("PID")))
        
        '记录已发布的菜单结点信息
        If .Cell(flexcpData, lngRow, .ColIndex("新")) = Val("0-已发布") Then
            Call RevokeAdd(lngRow)
        End If
        
        '删除
        .Redraw = False
        .RemoveItem .SelectedRow(0)
        If lngRow < .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        .Redraw = True
        .SetFocus
    End With
    
    '更新可选位置
    With tvwSelecting
        For lngRow = 1 To .Nodes.count
            If Val(.Nodes(lngRow).Tag) = lngID Then
                Call CheckGranted(lngID, .Nodes(lngRow))
                Exit For
            End If
        Next
    End With
End Sub

Private Sub cmdSingleSelect_Click()
'注：非末梢结点选中发布，该结点的子结点就有问题。因此，不支持选中后清除处理
    Dim objItem As ListItem
    Dim i As Long, lngRow As Long
    Dim blnFound As Boolean
    Dim strKey As String, strProgID As String, strText As String
    
    If cmdSingleSelect.Enabled = False Then Exit Sub
    
    If mblnChanged = False Then mblnChanged = True
    
    With vsfSelected
        For Each objItem In lvwReports.ListItems
            blnFound = False
            For i = 1 To .Rows - 1
                '排除界面“发布位置”列表的位置（报表ID 和 发布的菜单位置ID）
                If Val(.TextMatrix(i, .ColIndex("ReportID"))) = Val(objItem.Tag) _
                    And Val(.TextMatrix(i, .ColIndex("PID"))) = Val(tvwSelecting.SelectedItem.Tag) Then
                    blnFound = True
                    Exit For
                End If
            Next
            
            If blnFound = False Then
                '新增
                .Redraw = False
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                
                '已选标识
                tvwSelecting.SelectedItem.Bold = True
                
                '排除已发布的项目（可能被撤销中）
                strKey = CStr(Val(tvwSelecting.SelectedItem.Tag)) & "_" & Val(objItem.Tag)    '菜单ID_报表ID/组ID
                If CollectionFind(mcolRevoke, strKey) Then
                    .TextMatrix(lngRow, .ColIndex("程序ID")) = mcolRevoke(strKey).ReadProperty("ProgID")
                    .TextMatrix(lngRow, .ColIndex("PID")) = mcolRevoke(strKey).ReadProperty("PID")
                    .TextMatrix(lngRow, .ColIndex("SysNo")) = mcolRevoke(strKey).ReadProperty("SysNo")
                    Call CollectionDelete(mcolRevoke, strKey)
                Else
                    If Mode_ = 导航台 Then
                        strProgID = Trim$(Mid$(objItem.Tag, InStr(objItem.Tag, "_") + 1))
                    Else
                        strProgID = Val(tvwSelecting.SelectedItem.Tag)
                    End If
                    .TextMatrix(lngRow, .ColIndex("程序ID")) = strProgID
                    .TextMatrix(lngRow, .ColIndex("PID")) = Val(tvwSelecting.SelectedItem.Tag)
                    .TextMatrix(lngRow, .ColIndex("SysNo")) = Abs(Val(GetRootNode(tvwSelecting.SelectedItem).Tag))
                    .Cell(flexcpData, lngRow, .ColIndex("新")) = Val("1-新增")
                    .Cell(flexcpPicture, lngRow, .ColIndex("新")) = img16.ListImages("NEW").Picture
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("新")) = flexPicAlignCenterCenter
                End If
                strText = GetRootNode(tvwSelecting.SelectedItem).Text
                If InStr(strText, "]") > 0 Then
                    .TextMatrix(lngRow, .ColIndex("系统")) = Mid$(strText, InStr(strText, "]") + 1)
                Else
                    .TextMatrix(lngRow, .ColIndex("系统")) = strText
                End If
                .TextMatrix(lngRow, .ColIndex("组别")) = cboMenuGroup.Text
                .TextMatrix(lngRow, .ColIndex("菜单")) = tvwSelecting.SelectedItem.Text
                .TextMatrix(lngRow, .ColIndex("报表")) = objItem.Text
                .TextMatrix(lngRow, .ColIndex("MenuID")) = ""
                .TextMatrix(lngRow, .ColIndex("ReportID")) = CStr(Val(objItem.Tag))
                .Row = lngRow
                .TopRow = .BottomRow
                .Redraw = True
            End If
        Next
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Call FindItem(False)
    End If
End Sub

Private Sub Form_Load()
    mblnResult = False
    mblnChanged = False
    Me.KeyPreview = True
End Sub

Private Sub RefreshSelected()
    Dim strSQL As String, strReportID As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo hErr
    
    strReportID = GetReportIDs()
    
    If Mode_ = 导航台 Then
        '报表已发布的位置
        If mblnGroup Then
            strSQL = _
                "Select a2.组别, a2.ID PID, a2.标题 菜单, b.名称 系统, a1.ID MenuID, d.名称 报表, d.ID ReportID" & vbCr & _
                "  , d.程序id " & vbCr & _
                "From zlMenus A1, zlMenus A2, zlSystems B, zlPrograms C, zlRPTGroups D, Table(f_Num2list([1], ',')) E " & vbCr & _
                "Where a1.模块 = c.序号 And a1.模块 = d.程序id And a2.系统 = b.编号 " & vbCr & _
                "  And a1.上级id = a2.ID(+) and d.ID = e.Column_Value " & vbCr & _
                "  And Upper(c.部件) = 'ZL9REPORT' " & vbCr & _
                "  And a1.系统 is Null And c.系统 is Null And d.系统 is Null "
        Else
            strSQL = _
                "Select a2.组别, a2.ID PID, a2.标题 菜单, b.名称 系统, a1.ID MenuID, d.名称 报表, d.ID ReportID" & vbCr & _
                "  , d.程序id " & vbCr & _
                "From zlMenus A1, zlMenus A2, zlSystems B, zlPrograms C, zlReports D, Table(f_Num2list([1], ',')) E " & vbCr & _
                "Where a1.模块 = c.序号 And a1.模块 = d.程序id And a2.系统 = b.编号 " & vbCr & _
                "  And a1.上级id = a2.ID(+) And d.ID = e.Column_Value " & vbCr & _
                "  And Upper(c.部件) = 'ZL9REPORT' " & vbCr & _
                "  And a1.系统 is Null And c.系统 is Null And d.系统 is Null "
        End If
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取报表已发布的菜单位置", strReportID)
    Else
        strSQL = _
            "Select a.组别, a.模块 Pid, a.标题 菜单, b.名称 系统, a.Id Menuid, e.名称 报表, e.Id ReportID" & vbCr & _
            "  , a.模块 程序id, a.系统 SysNo " & vbCr & _
            "From zlMenus A, zlSystems B, zlPrograms C, zlRPTPuts D, zlReports E " & vbCr & _
            "  , (Select 系统 * 100 系统, 序号 From zlRegFunc Group By 系统, 序号) F " & vbCr & _
            "  , Table(Cast(f_Num2list([1], ',') As t_Numlist)) G " & vbCr & _
            "Where a.系统 = b.编号 " & vbCr & _
            "  And a.模块 = c.序号 " & vbCr & _
            "  And c.系统 = d.系统 And c.序号 = d.程序id " & vbCr & _
            "  And c.系统 = f.系统 And c.序号 = f.序号 " & vbCr & _
            "  And d.报表id = e.Id " & vbCr & _
            "  And e.Id = g.Column_Value " & vbCr & _
            "  And Upper(c.部件) <> 'ZL9REPORT' "
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取报表已发布的模块位置", strReportID)
    End If
    
    mobjSelected.Recordset = rsTemp
    Call mobjSelected.Repaint(RT_Rows)
    rsTemp.Close
    
    Exit Sub
    
hErr:
    If mdlPublic.ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshMenuGroup()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objNode As Node
    Dim bytStep As Byte

    On Error GoTo hErr
    
    bytStep = Val("0-步骤")
    
    If Mode_ = 导航台 Then
        strSQL = _
            "Select * From (" & vbCr & _
            "  Select 编号 As Scol, 0 As Flag, -编号 As ID, -null As 上级id, '[' || 编号 || ']' || 名称 As 标题 " & vbCr & _
            "  From zlSystems A " & vbCr & _
            "  Where Exists(Select 1 From zlMenus B Where b.系统 = a.编号 And b.组别 = [1]) " & vbCr & _
            "  Union All " & vbCr & _
            "  Select 99999 As Scol, Level As Flag, ID, Nvl(上级id, -系统) As 上级id, 标题 " & vbCr & _
            "  From zlMenus A " & vbCr & _
            "  Where 组别 = [1] And 模块 Is Null " & vbCr & _
            "    And Exists(Select 1 From zlSystems B Where b.编号 = a.系统) " & vbCr & _
            "  Start With 上级id Is Null And 组别 = [1] " & vbCr & _
            "  Connect By Prior ID = 上级id And 组别 = [1] " & vbCr & _
            ") Order By Scol, Flag, ID"
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取发布至导航台的菜单树", cboMenuGroup.Text)
    Else
        strSQL = _
            "Select * From (" & vbCr & _
            "  Select 0 Flag, -null 上级id, -编号 ID, '[' || 编号 || ']' || 名称 标题 " & vbCr & _
            "  From zlSystems A " & vbCr & _
            "  Where Exists(Select 1 From zlMenus B Where b.系统 = a.编号 And b.组别 = [1]) " & vbCr & _
            "  Union All " & vbCr & _
            "  Select Distinct 1 Flag, -b.系统 上级id, b.序号 ID, b.标题 " & vbCr & _
            "  From zlMenus A, zlPrograms B " & vbCr & _
            "    , (Select 系统 * 100 系统, 序号 From zlRegFunc Group By 系统, 序号) C " & vbCr & _
            "  Where a.系统 = b.系统 And a.模块 = b.序号 " & vbCr & _
            "    And b.系统 = c.系统 And b.序号 = c.序号 " & vbCr & _
            "    And Upper(b.部件) <> 'ZL9REPORT' And a.组别 = [1] " & vbCr & _
            "    And Exists(Select 1 From zlSystems X Where x.编号 = a.系统) " & vbCr & _
            ") Order By Flag, Abs(上级ID), Abs(id) "
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取发布至模块的菜单树", cboMenuGroup.Text)
    End If

    bytStep = Val("1-步骤")
    
    With tvwSelecting
        .Nodes.Clear
        Do While rsTemp.EOF = False
            If Nvl(rsTemp!Flag, 0) = 0 Then
                Set objNode = .Nodes.Add(, , "_" & rsTemp!id, rsTemp!标题, "SYS")
            Else
                If Mode_ = 导航台 Then
                    Set objNode = .Nodes.Add("_" & rsTemp!上级ID, tvwChild _
                        , "_" & rsTemp!id, rsTemp!标题, "NODE")
                Else
                    Set objNode = .Nodes.Add("_" & rsTemp!上级ID, tvwChild _
                        , "_" & rsTemp!id & "_" & Nvl(rsTemp!上级ID), rsTemp!标题, "MODULE")
                End If
            End If
            objNode.Tag = rsTemp!id
            objNode.Expanded = True
            Call CheckGranted(rsTemp!id, objNode)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End With
    
    Exit Sub
    
hErr:
    Select Case bytStep
    Case 1
        MsgBox "提醒：" & vbCr & "菜单结构可能存在异常，请检查后台数据！", vbInformation, App.Title
    Case Else
        If mdlPublic.ErrCenter() = 1 Then Resume
    End Select
End Sub

Private Sub CheckGranted(ByVal lngID As Long, ByRef objNode As Node)
    Dim l As Long
    
    With vsfSelected
        For l = 1 To .Rows - 1
            If lngID = Val(.TextMatrix(l, .ColIndex("PID"))) Then
                objNode.Bold = True
                Exit Sub
            End If
        Next
        objNode.Bold = False
    End With
End Sub

Private Sub InitSelected()
    Set mobjSelected = New clsVSFlexGridEx
    With mobjSelected
        .AppTemplate EM_Display, vsfSelected, MSTR_NAVIGATION, "", True
        .Init False
        .Binding.ExplorerBar = flexExNone
    End With
    
    If lblReportName.Visible Then
        With vsfSelected
            .ColHidden(.ColIndex("报表")) = True
            .ColWidth(.ColIndex("菜单")) = .ColWidth(.ColIndex("菜单")) + .ColWidth(.ColIndex("报表"))
        End With
    End If
End Sub

Private Sub InitSelecting()
    With tvwSelecting
        .Appearance = ccFlat
        .BorderStyle = ccFixedSingle
        .FullRowSelect = True
        .Indentation = 300
        .LineStyle = tvwRootLines
        Set .ImageList = img16
    End With
End Sub

Private Sub InitMenuGroup()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo hErr
    
    With cboMenuGroup
        .Appearance = 0
        .Clear
    End With
    
    strSQL = "Select Distinct 组别 From zlMenus Where 组别 Is Not Null "
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取菜单组别")
    Do While rsTemp.EOF = False
        cboMenuGroup.AddItem rsTemp!组别
        If rsTemp!组别 = "缺省" Then
            cboMenuGroup.ListIndex = cboMenuGroup.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Exit Sub
    
hErr:
    If mdlPublic.ErrCenter() = 1 Then Resume
End Sub

Private Sub InitReportList(ByRef vsfSelect As VSFlexGrid, ByRef lngCount As Long)
    Dim lngRow As Long, lngSelect As Long
    Dim objItem As ListItem

    lngCount = 0
    With lvwReports
        Set .Icons = img16
        Set .SmallIcons = img16
        .AllowColumnReorder = True
        .Appearance = ccFlat
        .BackColor = &H8000000F
        .View = lvwList
    End With
    
    With lvwReports.ColumnHeaders
        .Clear
        .Add , "ID"
        .Add , "Name"
    End With
    
    lngSelect = 0
    lvwReports.ListItems.Clear
    For lngRow = 1 To vsfSelect.Rows - 1
        If vsfSelect.SelectedRow(lngSelect) = lngRow Then
            If LCase$(vsfSelect.name) = "vsfgroup" Then
                Set objItem = lvwReports.ListItems.Add( _
                    , "_" & vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("ID")) _
                    , vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("组名")) _
                    , "REPORT", "REPORT")
            Else
                Set objItem = lvwReports.ListItems.Add( _
                    , "_" & vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("ID")) _
                    , vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("名称")) _
                    , "REPORT", "REPORT")
            End If
            objItem.Tag = vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("ID")) & _
                          "_" & _
                          vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("程序ID"))
            lngSelect = lngSelect + 1
        End If
    Next
    lngCount = lngSelect
    
    If lngCount = 1 Then
        lvwReports.Visible = False
        lblReportName.Visible = True
        lblReportName.Caption = lvwReports.ListItems(1).Text
    Else
        lvwReports.Visible = True
        lblReportName.Visible = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChanged And mblnResult = False Then
        If MsgBox("是否确定放弃发布？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If lvwReports.Visible = False Then
        With fraSelecting
            .Top = lblReportName.Top + lblReportName.Height + 150
            .Height = lblPos.Top - .Top - 30
        End With
        
        tvwSelecting.Height = fraSelecting.Height - tvwSelecting.Top - txtFind.Height - 210
        txtFind.Top = fraSelecting.Height - txtFind.Height - 120
        lblFind.Top = txtFind.Top + 30
                
        With fraSelectted
            .Top = fraSelecting.Top
            .Height = fraSelecting.Height
        End With
        
        vsfSelected.Height = fraSelectted.Height - vsfSelected.Top - txtFind.Height - 210
        cmdClear.Top = vsfSelected.Top + vsfSelected.Height + 45
        cmdSave.Top = cmdClear.Top
        cmdCancel.Top = cmdClear.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcolRevoke = Nothing
    Set mobjSelected = Nothing
End Sub

Private Sub lvwReports_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1      '禁止修改
End Sub

Private Function GetReportIDs()
    Dim i As Integer
    Dim strResult As String
    
    For i = 1 To lvwReports.ListItems.count
        strResult = strResult & "," & Mid$(lvwReports.ListItems(i).Key, 2)
    Next
    If strResult <> "" Then
        GetReportIDs = Mid$(strResult, 2)
    End If
End Function

Private Sub tvwSelecting_Click()
    If tvwSelecting.SelectedItem Is Nothing Then
        cmdSingleSelect.Enabled = False
    Else
        cmdSingleSelect.Enabled = Not tvwSelecting.SelectedItem.Parent Is Nothing
    End If
End Sub

Private Sub tvwSelecting_DblClick()
    If tvwSelecting.SelectedItem.Parent Is Nothing Then
        cmdSingleSelect.Enabled = False
    Else
        cmdSingleSelect.Enabled = True
        Call cmdSingleSelect_Click
        If tvwSelecting.SelectedItem.Expanded = False Then
            tvwSelecting.SelectedItem.Expanded = True
        End If
    End If
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        '查找
        KeyCode = 0
        Call FindItem(True)
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()+=[{]}\|;:'"",<.>/?", Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub vsfSelected_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    cmdSingleClear.Enabled = vsfSelected.Rows > 1
End Sub

Private Sub vsfSelected_DblClick()
    If vsfSelected.Row < 1 Then Exit Sub
    Call cmdSingleClear_Click
End Sub

Private Function GetRootNode(ByVal objNode As Node) As Node
    If objNode.Parent Is Nothing Then
        Set GetRootNode = objNode
    Else
        Set GetRootNode = GetRootNode(objNode.Parent)
    End If
End Function

Private Sub CollectionAdd(ByRef colVal As Collection, ByVal strKey As String, ByVal varVar As Variant)
    If colVal Is Nothing Then
        Set colVal = New Collection
    End If
    If CollectionFind(colVal, strKey) = False Then
        mcolRevoke.Add varVar, strKey
    End If
End Sub

Private Sub CollectionDelete(ByVal colVal As Collection, ByVal strKey As String)
    If CollectionFind(colVal, strKey) Then
        mcolRevoke.Remove strKey
    End If
End Sub

Private Function CollectionFind(ByVal colVal As Collection, ByVal strKey As String) As Boolean
    On Error Resume Next
    If IsObject(colVal.Item(strKey)) Then
        CollectionFind = Not colVal.Item(strKey) Is Nothing
    Else
        CollectionFind = colVal.Item(strKey) <> ""
    End If
    On Error GoTo 0
End Function

Private Function ReportVerify(ByVal lngReportID As Long, ByVal strName As String) As Boolean
    ReportVerify = False
    
    '验证密码
    If CheckPass(lngReportID) = False Then
        MsgBox mdlPublic.FormatString("【[1]】报表验证不能通过，拒绝发布！", strName) _
            , vbInformation, App.Title
        Exit Function
    End If
    
    '权限
    If CheckReportPriv(lngReportID) = False Then
        MsgBox mdlPublic.FormatString("你没有【[1]】报表中数据源涉及数据库对象的查询权限，请检查！", strName) _
            , vbInformation, App.Title
        Exit Function
    End If
    
    ReportVerify = True
End Function

Private Function GetProgID(ByVal lngReportID As Long, Optional blnGroup As Boolean = False) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    GetProgID = 0
    If blnGroup Then
        strSQL = "Select 程序ID from zlRPTGroups Where ID = [1]"
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取组报表的程序ID", lngReportID)
    Else
        strSQL = "Select 程序ID from zlReports Where ID = [1]"
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取独立、子报表的程序ID", lngReportID)
    End If
    If rsTemp.RecordCount > 0 Then
        GetProgID = mdlPublic.Nvl(rsTemp!程序id, 0)
    End If
    rsTemp.Close
    Exit Function
     
hErr:
    If mdlPublic.ErrCenter = 1 Then Resume
End Function

Private Function GetNodeNext(ByVal objNode As Node) As Node
    If objNode Is Nothing Then
        Exit Function
    Else
        If Not objNode.Next Is Nothing Then
            Set GetNodeNext = objNode.Next
        Else
            Set GetNodeNext = GetNodeNext(objNode.Parent)
        End If
    End If
End Function

Private Function FindItemRecursive(ByVal strFind As String, ByVal objNode As Node) As Node
    If objNode Is Nothing Then
        Exit Function
    End If

    If UCase$(objNode.Text) Like "*" & UCase$(Trim$(strFind)) & "*" And mlngPrevious < objNode.Index Then
        Set FindItemRecursive = objNode
        mlngPrevious = objNode.Index
    Else
        If Not objNode.Child Is Nothing Then
            Set FindItemRecursive = FindItemRecursive(strFind, objNode.Child)
        ElseIf Not objNode.Next Is Nothing Then
            Set FindItemRecursive = FindItemRecursive(strFind, objNode.Next)
        Else
            Set FindItemRecursive = FindItemRecursive(strFind, GetNodeNext(objNode.Parent))
        End If
    End If
End Function

Private Sub FindItem(ByVal blnFirst As Boolean)
    Dim objFind As Node

    If blnFirst Then
        '首次查找
        mlngPrevious = 1
    End If
    
    Set objFind = FindItemRecursive(txtFind.Text, tvwSelecting.Nodes(mlngPrevious))
    If Not objFind Is Nothing Then
        If Not tvwSelecting.SelectedItem Is Nothing Then
            If tvwSelecting.SelectedItem.Index = objFind.Index Then
                tvwSelecting.Nodes(1).Selected = True
            End If
        End If
        objFind.Selected = True
    Else
        If MsgBox("未查找到匹配的结点，是否从头开始查找？", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
            Call FindItem(True)
        End If
    End If
End Sub
