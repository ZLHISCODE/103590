VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmServiceList 
   BorderStyle     =   0  'None
   Caption         =   "frmServiceList"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picApp 
      BorderStyle     =   0  'None
      Height          =   3630
      Left            =   5010
      ScaleHeight     =   3630
      ScaleWidth      =   6090
      TabIndex        =   3
      Top             =   2415
      Width           =   6090
      Begin VB.PictureBox picImgApp 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   10
         Top             =   75
         Width           =   210
         Begin VB.Image imgColApp 
            Height          =   195
            Left            =   0
            Picture         =   "frmServiceList.frx":0000
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.CommandButton cmdDelApp 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   15
         MaskColor       =   &H8000000F&
         Picture         =   "frmServiceList.frx":054E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   465
         Visible         =   0   'False
         Width           =   265
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfApp 
         Height          =   2415
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3360
         _cx             =   5927
         _cy             =   4260
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   15658734
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmServiceList.frx":6DA0
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
      End
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   3810
      Left            =   1515
      ScaleHeight     =   3810
      ScaleWidth      =   7560
      TabIndex        =   1
      Top             =   915
      Width           =   7560
      Begin VB.CommandButton cmdChange 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   1110
         MaskColor       =   &H8000000F&
         Picture         =   "frmServiceList.frx":6DDC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1590
         Visible         =   0   'False
         Width           =   265
      End
      Begin VB.CommandButton cmdDel 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   660
         MaskColor       =   &H8000000F&
         Picture         =   "frmServiceList.frx":7366
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1590
         Visible         =   0   'False
         Width           =   265
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2415
         Left            =   0
         TabIndex        =   2
         Top             =   315
         Width           =   3360
         _cx             =   5927
         _cy             =   4260
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   15658734
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmServiceList.frx":DBB8
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   10
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.CheckBox chkAll 
            Height          =   180
            Left            =   60
            MaskColor       =   &H8000000F&
            TabIndex        =   9
            Top             =   285
            Width           =   210
         End
      End
      Begin VB.Image imgDel 
         Height          =   240
         Left            =   4125
         Picture         =   "frmServiceList.frx":DCA1
         Top             =   1950
         Width           =   240
      End
      Begin VB.Image imgDoc 
         Height          =   240
         Left            =   3900
         Picture         =   "frmServiceList.frx":144F3
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "停诊时间"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   2670
         TabIndex        =   8
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "停诊号码"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   60
         Width           =   720
      End
   End
   Begin XtremeSuiteControls.TabControl tabMain 
      Height          =   1950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3300
      _Version        =   589884
      _ExtentX        =   5821
      _ExtentY        =   3440
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList img16 
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceList.frx":14A7D
            Key             =   "K1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceList.frx":15017
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imdDoc 
      Height          =   240
      Left            =   3945
      Picture         =   "frmServiceList.frx":155B1
      Top             =   435
      Width           =   240
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   705
      Top             =   4065
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmServiceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mcbrPopupMain As CommandBar
Private mblnMark As Boolean
Public mfrmApp As New frmServiceApp
Public mfrmAppHistory As New frmServiceHistory
Private mfrmMain As Object
Private mlng消息ID As Long
Private mstrPreNote As String

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmMain.ExecuteCommandBars(Control)
End Sub

Public Sub InitFrmMain(frmMain As Object, strPrivs As String)
    Set mfrmMain = frmMain
    mstrPrivs = strPrivs
End Sub

Private Sub DeleteRecord()
    Dim strNO As String, strSQL As String, rsTemp As ADODB.Recordset
    Dim bln预约 As Boolean, lngRow As Long
    If tabMain(0).Selected Then
        lngRow = vsfList.Row
        strNO = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("单据号"))
    Else
        strNO = vsfApp.TextMatrix(vsfApp.Row, vsfApp.ColIndex("单据号"))
    End If
    strSQL = "Select 1 From 病人挂号记录 Where 记录性质=2 And No=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then
        bln预约 = False
    Else
        bln预约 = True
    End If
    strSQL = "Select 操作员姓名,登记时间 From 病人挂号记录 Where 记录状态=1 And No=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then
        MsgBox "没有找到预约记录,不能取消预约!", vbInformation, gstrSysName
        Exit Sub
    End If
    Call mfrmMain.DeleteRecord(strNO, Nvl(rsTemp!操作员姓名), Nvl(rsTemp!登记时间), bln预约)
    Call LoadData(mlng消息ID)
    If lngRow <> 0 Then Call LocateNextRecord(lngRow)
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnable As Boolean, i As Integer
    Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            If tabMain.Item(1).Selected Or tabMain.Item(2).Selected Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case 3839 '换诊
            If tabMain.Item(1).Selected Or tabMain.Item(2).Selected Then
                Control.Enabled = False
            Else
                If tabMain.Item(0).Selected Then
                    If vsfList.Cell(flexcpForeColor, vsfList.Row, 1) <> vbRed And vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("消息ID")) <> "" Then
                        Control.Enabled = True
                    Else
                        Control.Enabled = False
                    End If
                Else
                    Control.Enabled = False
                End If
            End If
        Case 3004 '取消预约
            If tabMain.Item(1).Selected Then
                    Control.Enabled = False
                Else
                    If tabMain.Item(0).Selected Then
                        If vsfList.Cell(flexcpForeColor, vsfList.Row, 1) <> vbRed And vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("消息ID")) <> "" Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    End If
                    If tabMain.Item(2).Selected Then
                        Control.Enabled = True
                    End If
                    If tabMain.Item(3).Selected Then
                        If vsfApp.Cell(flexcpForeColor, vsfApp.Row, 1) <> vbRed And vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("消息ID")) <> "" Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    End If
                End If
                
        Case 2601 '已通知患者
            If tabMain.Item(3).Selected Or tabMain.Item(2).Selected Then
                    Control.Enabled = False
                Else
                    If tabMain.Item(0).Selected Then
                        If vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("是否处理")) = 0 And vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("消息ID")) <> "" Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    Else
                        Control.Enabled = False
                    End If
                End If
        Case 3950 '替诊
            If InStr(tabMain.Selected.Caption, "替诊") = 0 Then
                Control.Enabled = False
            Else
                If vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("是否处理")) = 0 And vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("消息ID")) <> "" And vsfList.Cell(flexcpForeColor, vsfList.Row, 1) <> vbRed Then
                    Control.Enabled = True
                Else
                    Control.Enabled = False
                End If
            End If
        Case 3936 '全部替诊
            If InStr(tabMain.Selected.Caption, "替诊") = 0 Then
                Control.Enabled = False
            Else
                blnEnable = False
                For i = 2 To vsfList.Rows - 1
                    If vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("是否处理")) = 0 And vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("消息ID")) <> "" And vsfList.Cell(flexcpForeColor, vsfList.Row, 1) <> vbRed Then blnEnable = True
                Next i
                Control.Enabled = blnEnable
            End If
        Case conMenu_Manage_Bespeak '预约挂号
            If tabMain.Item(1).Selected Then
                    Control.Enabled = InStr(mstrPrivs, ";预约登记信息处理;") > 0
                Else
                    If InStr(mstrPrivs, ";预约挂号登记;") = 0 Then
                        Control.Visible = False
                    Else
                        gobjRegist.zlUpdateCommandBars Control
                    End If
                End If
    End Select
End Sub

Public Sub LocateNextRecord(ByVal lngRow As Long)
    Dim i As Integer, lngActRow As Long
    On Error GoTo errH
    With vsfList
        For i = lngRow To .Rows - 1
            If .Cell(flexcpForeColor, i, 0, i, 0) = vbBlack And .TextMatrix(i, 0) = 0 Then
                lngActRow = i
                Exit For
            End If
        Next i
        If lngActRow <> 0 Then
            .Select lngActRow, 1
        Else
            .Select lngRow, 1
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chkAll_Click()
    Dim i As Integer, blnOperatable As Boolean, strCommand As String
    Dim strMsgResult As String, blnCancelable As Boolean
    If mblnMark Then Exit Sub
    If InStr(mstrPrivs, ";停诊信息处理;") = 0 Then
        MsgBox "权限不足,不能进行批量处理!", vbInformation, gstrSysName
        chkAll.Value = 0
        Exit Sub
    End If
    blnOperatable = False
    For i = 2 To vsfList.Rows - 1
        If Val(vsfList.TextMatrix(i, 0)) = 0 And Val(vsfList.RowData(i)) <> 0 And vsfList.Cell(flexcpForeColor, i, 1, i, 1) <> vbRed Then blnOperatable = True
    Next i
    If Not blnOperatable Then
        chkAll.Value = 1
        Exit Sub
    End If
    blnCancelable = True
    For i = 2 To vsfList.Rows - 1
        If vsfList.Cell(flexcpForeColor, i, 1, i, 1) <> vbRed Then
            If Val(vsfList.TextMatrix(i, 14)) <> 0 Then
                blnCancelable = False
            Else
                If vsfList.TextMatrix(i, 14) = "" Then
                    If Val(vsfList.TextMatrix(i, 7)) <> 0 Then
                        blnCancelable = False
                    End If
                End If
            End If
        End If
    Next i
    If blnCancelable Then
        strCommand = "批量通知患者,批量取消预约,取消"
    Else
        strCommand = "批量通知患者,取消"
    End If
    strMsgResult = zlCommFun.ShowMsgbox(gstrSysName, "是否批量处理全部未处理的记录?", strCommand, Me, vbQuestion)
    If strMsgResult = "" Or strMsgResult = "取消" Then
    mblnMark = True
    chkAll.Value = 0
    mblnMark = False
    Exit Sub
    End If
    If strMsgResult = "批量通知患者" Then
        Call mfrmMain.BatchInform
    End If
    If strMsgResult = "批量取消预约" Then
        Call mfrmMain.BatchCancel
    End If
    chkAll.Value = 1
End Sub

Private Sub cmdChange_Click()
    Call SelectDoc
End Sub

Private Sub cmdDel_Click()
    Call DeleteRecord
End Sub

Private Sub cmdDelApp_Click()
    Call DeleteRecord
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Call DefMainCommandBars
    tabMain.PaintManager.Appearance = xtpTabAppearancePropertyPage2003
    tabMain.PaintManager.BoldSelected = True
    tabMain.PaintManager.StaticFrame = True
    tabMain.PaintManager.ClientFrame = xtpTabFrameBorder
    tabMain.InsertItem 1, "停诊号别预约清单", picList.Hwnd, 0
    tabMain.InsertItem 2, "预约详细信息", mfrmApp.Hwnd, 0
    tabMain.InsertItem 3, "预约详细信息", mfrmAppHistory.Hwnd, 0
    tabMain.InsertItem 4, "历史预约信息", picApp.Hwnd, 0
    tabMain.Item(3).Selected = True
    tabMain.Item(0).Visible = False
    tabMain.Item(1).Visible = False
    tabMain.Item(2).Visible = False
    Call InitGrid
    Call LoadHistoryData
End Sub

Public Sub LoadHistoryData()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim dtpBegin As Date, dtpEnd As Date
    If mfrmMain Is Nothing Then Exit Sub
    If Format(mfrmMain.mdatBegin, "yyyy-mm-dd hh:mm:ss") = "1899-12-30 00:00:00" Then
        dtpBegin = zlDatabase.Currentdate - 365
        dtpEnd = zlDatabase.Currentdate
    Else
        dtpBegin = mfrmMain.mdatBegin
        dtpEnd = mfrmMain.mdatEnd
    End If
    strSQL = "" & vbNewLine & _
            " Select a.No As 单据号, a.预约时间, a.号别, a.号序, c.名称 As 科室, a.医生, e.名称 As 项目, a.门诊号, a.姓名, a.性别, d.身份证号, d.家庭电话 As 联系电话, a.费别," & vbNewLine & _
            "       a.摘要, a.挂号员, a.登记时间, a.记录状态, a.金额" & vbNewLine & _
            " From (Select b.No, Max(Nvl(b.预约时间, b.发生时间)) As 预约时间, Max(b.执行部门id) As 执行部门id, Max(b.号别) As 号别, Max(b.号序) As 号序," & vbNewLine & _
            "              Max(b.执行人) As 医生, Sum(c.实收金额) As 金额, Max(b.门诊号) As 门诊号, Max(b.病人id) As 病人id, Max(b.姓名) As 姓名," & vbNewLine & _
            "              Max(c.费别) As 费别, Max(b.性别) As 性别, Max(b.摘要) As 摘要, Max(b.操作员姓名) As 挂号员, b.登记时间, b.记录状态," & vbNewLine & _
            "              Max(b.出诊记录id) As 出诊记录id" & vbNewLine & _
            "       From 病人服务信息记录 A, 病人挂号记录 B, 门诊费用记录 C" & vbNewLine & _
            "       Where a.登记时间 Between [1] And [2] And (a.登记人 = [3] Or b.操作员姓名 = [3]) And a.挂号id = b.Id And b.No = c.No And" & vbNewLine & _
            "             c.记录性质 = 4" & vbNewLine & _
            "       Group By b.No, b.登记时间, b.记录状态) A, 部门表 C, 病人信息 D, 收费项目目录 E, 临床出诊记录 F" & vbNewLine & _
            " Where a.执行部门id = c.Id And a.病人id = d.病人id(+) And a.出诊记录id = f.Id And f.项目id = e.Id"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtpBegin, dtpEnd, UserInfo.姓名)
    With vsfApp
        .Clear 1
        .Rows = 2
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("单据号")) = Nvl(rsTemp!单据号)
            .TextMatrix(.Rows - 1, .ColIndex("预约时间")) = Format(Nvl(rsTemp!预约时间), "yyyy-mm-dd hh:mm")
            .TextMatrix(.Rows - 1, .ColIndex("号别")) = Nvl(rsTemp!号别)
            .TextMatrix(.Rows - 1, .ColIndex("号序")) = Nvl(rsTemp!号序)
            .TextMatrix(.Rows - 1, .ColIndex("科室")) = Nvl(rsTemp!科室)
            .TextMatrix(.Rows - 1, .ColIndex("医生")) = Nvl(rsTemp!医生)
            .TextMatrix(.Rows - 1, .ColIndex("门诊号")) = Nvl(rsTemp!门诊号)
            .TextMatrix(.Rows - 1, .ColIndex("项目")) = Nvl(rsTemp!项目)
            .TextMatrix(.Rows - 1, .ColIndex("姓名")) = Nvl(rsTemp!姓名)
            .TextMatrix(.Rows - 1, .ColIndex("性别")) = Nvl(rsTemp!性别)
            .TextMatrix(.Rows - 1, .ColIndex("联系电话")) = Nvl(rsTemp!联系电话)
            .TextMatrix(.Rows - 1, .ColIndex("费别")) = Nvl(rsTemp!费别)
            .TextMatrix(.Rows - 1, .ColIndex("摘要")) = Nvl(rsTemp!摘要)
            .TextMatrix(.Rows - 1, .ColIndex("挂号员")) = Nvl(rsTemp!挂号员)
            .TextMatrix(.Rows - 1, .ColIndex("登记时间")) = Format(Nvl(rsTemp!登记时间), "yyyy-mm-dd hh:mm")
            .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format(Val(Nvl(rsTemp!金额)), "0.00")

            If Val(Nvl(rsTemp!记录状态)) = 3 Then
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
            Else
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlack
            End If
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        For i = 0 To .Rows - 1
            .RowHeight(i) = 350
        Next i
        If rsTemp.RecordCount = 0 Then
            .AutoSize 1, .Cols - 1
        Else
            .AutoSize 0, .Cols - 1
        End If
        zl_vsGrid_Para_Restore 1115, vsfApp, Me.Name, "vsfApp"
    End With
End Sub

Public Sub RefreshData()
    Call LoadData(mlng消息ID)
End Sub

Private Sub InitGrid()
    Dim i As Integer
    With vsfList
        .Cols = 17
        .Rows = 3
        .FixedRows = 2
        .MergeCells = flexMergeFixedOnly
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(0, 1) = "姓名"
        .TextMatrix(1, 1) = "姓名"
        .ColWidth(1) = 800
        .TextMatrix(0, 2) = "门诊号"
        .TextMatrix(1, 2) = "门诊号"
        .ColWidth(2) = 900
        .TextMatrix(0, 3) = "性别"
        .TextMatrix(1, 3) = "性别"
        .ColWidth(3) = 500
        .TextMatrix(0, 4) = "年龄"
        .TextMatrix(1, 4) = "年龄"
        .ColWidth(4) = 500
        .TextMatrix(0, 5) = "联系电话"
        .TextMatrix(1, 5) = "联系电话"
        .TextMatrix(0, 6) = "预约信息"
        .TextMatrix(1, 6) = "单据号"
        .ColWidth(6) = 850
        .TextMatrix(0, 7) = "预约信息"
        .TextMatrix(1, 7) = "挂号费"
        .ColWidth(7) = 600
        .TextMatrix(0, 8) = "预约信息"
        .TextMatrix(1, 8) = "预约方式"
        .ColWidth(8) = 800
        .TextMatrix(0, 9) = "预约信息"
        .TextMatrix(1, 9) = "预约时间"
        .ColWidth(9) = 1600
        .TextMatrix(0, 10) = "换诊信息"
        .TextMatrix(1, 10) = "号码"
        .ColWidth(10) = 700
        .TextMatrix(0, 11) = "换诊信息"
        .TextMatrix(1, 11) = "项目"
        .TextMatrix(0, 12) = "换诊信息"
        .TextMatrix(1, 12) = "医生"
        .TextMatrix(0, 13) = "换诊信息"
        .TextMatrix(1, 13) = "科室"
        .TextMatrix(0, 14) = "换诊信息"
        .TextMatrix(1, 14) = "挂号费"
        .TextMatrix(0, 15) = "换诊信息"
        .TextMatrix(1, 15) = "预约时间"
        .TextMatrix(0, 16) = "消息ID"
        .TextMatrix(1, 16) = "消息ID"
        .MergeRow(0) = True
        For i = 0 To .Cols - 1
            If i = 0 Then
                .ColKey(i) = "是否处理"
            ElseIf i = 7 Then
                .ColKey(i) = "原挂号费"
            ElseIf i = 9 Then
                .ColKey(i) = "原预约时间"
            ElseIf i = 14 Then
                .ColKey(i) = "现挂号费"
            ElseIf i = 15 Then
                .ColKey(i) = "现预约时间"
            Else
                .ColKey(i) = .TextMatrix(1, i)
            End If
            If .ColKey(i) = "消息ID" Then .ColHidden(i) = True
            .MergeCol(i) = True
            .ColAlignment(i) = flexAlignCenterCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .TextMatrix(0, i) Like "*金额*" Then .ColAlignment(i) = flexAlignRightCenter
            If .TextMatrix(1, i) Like "*预约时间" Then .ColAlignment(i) = flexAlignLeftCenter
            If .TextMatrix(1, i) Like "*挂号费" Then .ColAlignment(i) = flexAlignRightCenter
            If .TextMatrix(1, i) = "门诊号" Then .ColAlignment(i) = flexAlignLeftCenter
            If .TextMatrix(1, i) = "联系电话" Then .ColAlignment(i) = flexAlignLeftCenter
            If .TextMatrix(1, i) = "科室" Then .ColAlignment(i) = flexAlignLeftCenter
        Next i
        For i = 0 To .Rows - 1
            .RowHeight(i) = 322
        Next i
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 11, .Cols - 1
        For i = 0 To .Cols - 1
            If i = 0 Then
                .ColWidth(i) = 350
            End If
        Next i
    End With
    
    With vsfApp
        .FixedRows = 1
        .Cols = 16
        .Rows = 2
        .TextMatrix(0, 0) = "预约时间"
        .ColData(0) = "-1|1"
        .ColWidth(0) = 1200
        .TextMatrix(0, 1) = "单据号"
        .ColData(1) = "-1|1"
        .TextMatrix(0, 2) = "科室"
        .TextMatrix(0, 3) = "医生"
        .TextMatrix(0, 4) = "项目"
        .TextMatrix(0, 5) = "金额"
        .ColData(5) = "-1|1"
        .TextMatrix(0, 6) = "费别"
        .TextMatrix(0, 7) = "姓名"
        .ColData(7) = "-1|1"
        .TextMatrix(0, 8) = "性别"
        .TextMatrix(0, 9) = "门诊号"
        .ColHidden(9) = True
        .TextMatrix(0, 10) = "联系电话"
        .TextMatrix(0, 11) = "摘要"
        .ColHidden(11) = True
        .TextMatrix(0, 12) = "号别"
        .ColHidden(12) = True
        .TextMatrix(0, 13) = "号序"
        .ColHidden(13) = True
        .TextMatrix(0, 14) = "挂号员"
        .TextMatrix(0, 15) = "登记时间"
        
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            If .ColKey(i) = "消息ID" Then .ColHidden(i) = True
            .MergeCol(i) = True
            .ColAlignment(i) = flexAlignCenterCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .TextMatrix(0, i) Like "*金额*" Then .ColAlignment(i) = flexAlignRightCenter
        Next i
        For i = 0 To .Rows - 1
            .RowHeight(i) = 322
        Next i
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 1, .Cols - 1
    End With
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    
    Err = 0: On Error GoTo errHandle
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 '去除菜单栏前缀
    
    Set mcbrPopupMain = cbsThis.Add("弹出菜单1", xtpBarPopup)
    With mcbrPopupMain.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&E)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "预约挂号(&A)"): cbrControl.BeginGroup = True
        cbrControl.IconId = 3003
        If InStr(mstrPrivs, ";预约挂号登记;") = 0 Then cbrControl.Visible = False
        If gbytRegistMode = 0 Then
            cbrControl.Visible = False
        Else
            If gdatRegistTime < zlDatabase.Currentdate Then
                cbrControl.Visible = False
            End If
        End If
        Set cbrControl = .Add(xtpControlButton, 3004, "取消预约(&C)")
        If InStr(mstrPrivs, ";停诊信息处理;") = 0 And InStr(mstrPrivs, ";预约登记信息处理;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 3839, "换诊(&H)")
        If InStr(mstrPrivs, ";停诊信息处理;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 3950, "替诊(&T)")
        If InStr(mstrPrivs, ";停诊信息处理;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 3936, "全部替诊(&O)")
        If InStr(mstrPrivs, ";停诊信息处理;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 2601, "已通知患者(&Z)")
        cbrControl.IconId = 11151
        If InStr(mstrPrivs, ";停诊信息处理;") = 0 And InStr(mstrPrivs, ";预约登记信息处理;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    If InStr(mstrPrivs, ";停诊信息处理;") = 0 Then chkAll.Enabled = False
    
    DefMainCommandBars = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub LoadData(ByVal lngID As Long)
    Dim i As Integer, rsTemp As ADODB.Recordset, rsItem As ADODB.Recordset
    Dim rsDetail As ADODB.Recordset, lng记录ID As Long
    Dim strSQL As String, intType As Integer, lngRepeated As Long
    Dim lngTotal As Long, lngProcessed As Long, lngCanceled As Long, lngChanged As Long
    Dim strPriceGrade As String
    
    On Error GoTo errHandle
    cmdDel.Visible = False
    cmdChange.Visible = False
    mlng消息ID = lngID
    strSQL = "Select 通知类型,处理时间,记录ID From 病人服务信息记录 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    intType = Val(Nvl(rsTemp!通知类型))
    lng记录ID = Val(Nvl(rsTemp!记录ID))
    Select Case intType
    Case 1
        mblnMark = True
        chkAll.Value = 0
        mblnMark = False
        tabMain.Item(0).Caption = "预约清单(停诊)"
        tabMain.Item(2).Visible = False
        tabMain.Item(1).Visible = False
        tabMain.Item(0).Visible = True
        tabMain.Item(0).Selected = True
        strSQL = "Select b.号码, b.号类, c.名称 As 科室, d.名称 As 项目, Nvl(e.姓名, a.医生姓名) As 医生, f.开始时间, f.终止时间 " & vbNewLine & _
                "From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 收费项目目录 D, 人员表 E, 临床出诊停诊记录 F" & vbNewLine & _
                "Where a.Id = [1] And a.号源id = b.Id And b.科室id = c.Id And a.项目id = d.Id And a.医生id = e.Id(+) And a.id = f.记录id(+) And f.取消时间 Is Null"
        Set rsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
        lblInfo.Caption = "停诊号码:" & rsDetail!号码 & "   "
        lblInfo.Caption = lblInfo.Caption & "停诊号类:" & rsDetail!号类 & "   "
        lblInfo.Caption = lblInfo.Caption & "停诊科室:" & rsDetail!科室 & "   "
        lblInfo.Caption = lblInfo.Caption & "停诊项目:" & rsDetail!项目 & "   "
        lblInfo.Caption = lblInfo.Caption & "停诊医生:" & rsDetail!医生 & "   "
        lblTime.Caption = "停诊时间:" & Format(rsDetail!开始时间, "yyyy-mm-dd hh:mm:ss") & " 至 " & Format(rsDetail!终止时间, "yyyy-mm-dd hh:mm:ss")
        lblTime.Left = lblInfo.Left + lblInfo.Width
        strSQL = "Select a.处理时间, b.病人ID, b.医疗付款方式, c.姓名, c.门诊号, c.性别, c.年龄, b.家庭电话, c.No As 单据号, c.预约方式, Nvl(d.原预约时间, Nvl(c.预约时间,c.发生时间)) As 原预约时间," & vbNewLine & _
                "       Nvl(d.原项目id, e.项目id) As 原项目id, d.现号码, d.现项目id, d.现医生姓名, f.名称 As 现科室, d.现预约时间, C.记录状态, a.id As 消息ID, g.名称 As 换诊项目" & vbNewLine & _
                "From 病人服务信息记录 A, 病人信息 B, 病人挂号记录 C, 就诊变动记录 D, 临床出诊记录 E, 部门表 F, 收费项目目录 G" & vbNewLine & _
                "Where a.记录id = e.Id And a.通知类型 = 1 And a.记录id = [1] And a.病人id = b.病人id(+) And a.挂号id = c.Id And d.现项目id = g.Id(+) And d.挂号单(+) = c.No And" & vbNewLine & _
                "      d.现科室id = f.Id(+) And (d.登记时间 = (Select Max(登记时间) From 就诊变动记录 Where 挂号单 = c.No) Or d.登记时间 Is Null)  "
        Set rsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
        With vsfList
            .Clear 1
            .Rows = 3
            Do While Not rsDetail.EOF
                lngTotal = lngTotal + 1
                .TextMatrix(.Rows - 1, 0) = IIf(IsNull(rsDetail!处理时间), 0, 1)
                If Not IsNull(rsDetail!处理时间) Then
                    lngProcessed = lngProcessed + 1
                    If Val(Nvl(rsDetail!记录状态)) = 3 Then
                        lngRepeated = lngRepeated + 1
                    End If
                End If
                '价格等级
                If gintPriceGradeStartType >= 2 Then
                    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(rsDetail!病人ID)), 0, Nvl(rsDetail!医疗付款方式, ""), , , strPriceGrade)
                Else
                    strPriceGrade = gstrPriceGrade
                End If
                .TextMatrix(.Rows - 1, 1) = Nvl(rsDetail!姓名)
                .Cell(flexcpData, .Rows - 1, 1) = strPriceGrade
                .TextMatrix(.Rows - 1, 2) = Nvl(rsDetail!门诊号)
                .TextMatrix(.Rows - 1, 3) = Nvl(rsDetail!性别)
                .TextMatrix(.Rows - 1, 4) = Nvl(rsDetail!年龄)
                .TextMatrix(.Rows - 1, 5) = Nvl(rsDetail!家庭电话)
                .TextMatrix(.Rows - 1, 6) = rsDetail!单据号
                .TextMatrix(.Rows - 1, 7) = FormatEx(Get项目金额(Val(Nvl(rsDetail!原项目ID)), strPriceGrade), 2)
                .TextMatrix(.Rows - 1, 8) = Nvl(rsDetail!预约方式)
                .TextMatrix(.Rows - 1, 9) = Format(Nvl(rsDetail!原预约时间), "yyyy-mm-dd hh:mm")
                .TextMatrix(.Rows - 1, 10) = rsDetail!现号码 & "  "
                .TextMatrix(.Rows - 1, 11) = Nvl(rsDetail!换诊项目)
                .TextMatrix(.Rows - 1, 12) = Nvl(rsDetail!现医生姓名)
                .TextMatrix(.Rows - 1, 13) = Nvl(rsDetail!现科室)
                If IsNull(rsDetail!现项目ID) Then
                    .TextMatrix(.Rows - 1, 14) = ""
                Else
                    .TextMatrix(.Rows - 1, 14) = FormatEx(Get项目金额(Val(Nvl(rsDetail!现项目ID)), strPriceGrade), 2)
                End If
                .TextMatrix(.Rows - 1, 15) = Format(Nvl(rsDetail!现预约时间), "yyyy-mm-dd hh:mm")
                .TextMatrix(.Rows - 1, 16) = Nvl(rsDetail!消息ID)
                .RowData(.Rows - 1) = Nvl(rsDetail!消息ID)
                If Val(Nvl(rsDetail!记录状态)) = 3 Then
                    lngCanceled = lngCanceled + 1
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                Else
                    If IsNull(rsDetail!现号码) Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlack
                    Else
                        lngChanged = lngChanged + 1
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
                    End If
                End If
                .Rows = .Rows + 1
                rsDetail.MoveNext
            Loop
            If .Rows > 3 Then .Rows = .Rows - 1
            For i = 0 To .Rows - 1
                .RowHeight(i) = 322
            Next i
            .Select 2, 1
            .AutoSize 11, .Cols - 1
            zl_vsGrid_Para_Restore 1115, vsfList, Me.Name, "vsfList"
            Call vsfList_EnterCell
            mfrmMain.stbThis.Panels(2).Text = "当前共有" & lngTotal & "条预约单,其中已经取消" & lngCanceled & "条,换诊" & lngChanged & "条,处理" & lngProcessed & "条"
            mblnMark = True
            If lngTotal = lngProcessed + lngCanceled - lngRepeated Then chkAll.Value = 1
            mblnMark = False
        End With
    Case 2
        mblnMark = True
        chkAll.Value = 0
        mblnMark = False
        tabMain.Item(0).Caption = "预约清单(替诊)"
        tabMain.Item(2).Visible = False
        tabMain.Item(1).Visible = False
        tabMain.Item(0).Visible = True
        tabMain.Item(0).Selected = True
        strSQL = "Select b.号码, b.号类, c.名称 As 科室, d.名称 As 项目, Nvl(e.姓名, a.医生姓名) As 医生, f.开始时间, f.终止时间, a.替诊医生姓名 As 替诊医生" & vbNewLine & _
                "From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 收费项目目录 D, 人员表 E, 临床出诊停诊记录 F" & vbNewLine & _
                "Where a.Id = [1] And a.号源id = b.Id And b.科室id = c.Id And a.项目id = d.Id And a.医生id = e.Id(+) And a.id = f.记录id(+) And f.取消时间 Is Null"
        Set rsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
        lblInfo.Caption = "停诊号码:" & rsDetail!号码 & "   "
        lblInfo.Caption = lblInfo.Caption & "停诊号类:" & rsDetail!号类 & "   "
        lblInfo.Caption = lblInfo.Caption & "停诊科室:" & rsDetail!科室 & "   "
        lblInfo.Caption = lblInfo.Caption & "停诊项目:" & rsDetail!项目 & "   "
        lblInfo.Caption = lblInfo.Caption & "停诊医生:" & rsDetail!医生 & "   "
        lblInfo.Caption = lblInfo.Caption & "替诊医生:" & rsDetail!替诊医生 & "   "
        lblTime.Caption = "停诊时间:" & Format(rsDetail!开始时间, "yyyy-mm-dd hh:mm:ss") & " 至 " & Format(rsDetail!终止时间, "yyyy-mm-dd hh:mm:ss")
        lblTime.Left = lblInfo.Left + lblInfo.Width
        strSQL = "Select a.处理时间, b.病人ID, b.医疗付款方式, c.姓名, c.门诊号, c.性别, c.年龄, b.家庭电话, c.No As 单据号, c.预约方式, Nvl(d.原预约时间, Nvl(c.预约时间,c.发生时间)) As 原预约时间," & vbNewLine & _
                "       Nvl(d.原项目id, e.项目id) As 原项目id, d.现号码, d.现项目id, d.现医生姓名, f.名称 As 现科室, d.现预约时间, c.记录状态, a.id As 消息ID, g.名称 As 换诊项目" & vbNewLine & _
                "From 病人服务信息记录 A, 病人信息 B, 病人挂号记录 C, 就诊变动记录 D, 临床出诊记录 E, 部门表 F, 收费项目目录 G" & vbNewLine & _
                "Where a.记录id = e.Id And a.通知类型 = 2 And a.记录id = [1] And a.病人id = b.病人id(+) And d.现项目id = g.Id(+) And a.挂号id = c.Id And d.挂号单(+) = c.No And" & vbNewLine & _
                "      d.现科室id = f.Id(+) And (d.登记时间 = (Select Max(登记时间) From 就诊变动记录 Where 挂号单 = c.No) Or d.登记时间 Is Null) "
        
        Set rsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
        With vsfList
            .Clear 1
            .Rows = 3
            Do While Not rsDetail.EOF
                lngTotal = lngTotal + 1
                .TextMatrix(.Rows - 1, 0) = IIf(IsNull(rsDetail!处理时间), 0, 1)
                If Not IsNull(rsDetail!处理时间) Then
                    lngProcessed = lngProcessed + 1
                    If Val(Nvl(rsDetail!记录状态)) = 3 Then
                        lngRepeated = lngRepeated + 1
                    End If
                End If
                '价格等级
                If gintPriceGradeStartType >= 2 Then
                    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(rsDetail!病人ID)), 0, Nvl(rsDetail!医疗付款方式, ""), , , strPriceGrade)
                Else
                    strPriceGrade = gstrPriceGrade
                End If
                .TextMatrix(.Rows - 1, 1) = Nvl(rsDetail!姓名)
                .Cell(flexcpData, .Rows - 1, 1) = strPriceGrade
                .TextMatrix(.Rows - 1, 2) = Nvl(rsDetail!门诊号)
                .TextMatrix(.Rows - 1, 3) = Nvl(rsDetail!性别)
                .TextMatrix(.Rows - 1, 4) = Nvl(rsDetail!年龄)
                .TextMatrix(.Rows - 1, 5) = Nvl(rsDetail!家庭电话)
                .TextMatrix(.Rows - 1, 6) = rsDetail!单据号
                .TextMatrix(.Rows - 1, 7) = FormatEx(Get项目金额(Val(Nvl(rsDetail!原项目ID)), strPriceGrade), 2)
                .TextMatrix(.Rows - 1, 8) = Nvl(rsDetail!预约方式)
                .TextMatrix(.Rows - 1, 9) = Format(Nvl(rsDetail!原预约时间), "yyyy-mm-dd hh:mm")
                .TextMatrix(.Rows - 1, 10) = rsDetail!现号码 & "  "
                .TextMatrix(.Rows - 1, 11) = Nvl(rsDetail!换诊项目)
                .TextMatrix(.Rows - 1, 12) = Nvl(rsDetail!现医生姓名)
                .TextMatrix(.Rows - 1, 13) = Nvl(rsDetail!现科室)
                If IsNull(rsDetail!现项目ID) Then
                    .TextMatrix(.Rows - 1, 14) = ""
                Else
                    .TextMatrix(.Rows - 1, 14) = FormatEx(Get项目金额(Val(Nvl(rsDetail!现项目ID)), strPriceGrade), 2)
                End If
                .TextMatrix(.Rows - 1, 15) = Format(Nvl(rsDetail!现预约时间), "yyyy-mm-dd hh:mm")
                .TextMatrix(.Rows - 1, 16) = Nvl(rsDetail!消息ID)
                .RowData(.Rows - 1) = Nvl(rsDetail!消息ID)
                If Val(Nvl(rsDetail!记录状态)) = 3 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                    lngCanceled = lngCanceled + 1
                Else
                    If IsNull(rsDetail!现号码) Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlack
                    Else
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
                        lngChanged = lngChanged + 1
                    End If
                End If
                .Rows = .Rows + 1
                rsDetail.MoveNext
            Loop
            If .Rows > 3 Then .Rows = .Rows - 1
            For i = 0 To .Rows - 1
                .RowHeight(i) = 322
            Next i
            .Select 2, 1
            .AutoSize 11, .Cols - 1
            zl_vsGrid_Para_Restore 1115, vsfList, Me.Name, "vsfList"
            Call vsfList_EnterCell
            mfrmMain.stbThis.Panels(2).Text = "当前共有" & lngTotal & "条预约单,其中已经取消" & lngCanceled & "条,换诊" & lngChanged & "条,处理" & lngProcessed & "条"
            mblnMark = True
            If lngTotal = lngProcessed + lngCanceled - lngRepeated Then chkAll.Value = 1
            mblnMark = False
        End With
    Case 3
        If IsNull(rsTemp!处理时间) Then
            '未处理的预约登记
            tabMain.Item(2).Visible = False
            tabMain.Item(1).Visible = True
            tabMain.Item(0).Visible = False
            tabMain.Item(1).Selected = True
            Call mfrmApp.LoadData(Me, lngID)
        Else
            '已处理的预约登记
            tabMain.Item(2).Visible = True
            tabMain.Item(1).Visible = False
            tabMain.Item(0).Visible = False
            tabMain.Item(2).Selected = True
            Call mfrmAppHistory.LoadData(lngID)
        End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub ShowPanelText(ByVal strText As String)
    mfrmMain.stbThis.Panels(2).Text = strText
End Sub

Public Function DirectApp() As Boolean
    If tabMain.Item(1).Selected = False Then Exit Function
    If mfrmApp.SaveData = False Then Exit Function
    DirectApp = True
End Function

Private Sub Form_Resize()
    With tabMain
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save 1115, vsfList, Me.Name, "vsfList"
    zl_vsGrid_Para_Save 1115, vsfApp, Me.Name, "vsfApp"
    Unload mfrmApp
    Unload mfrmAppHistory
End Sub

Private Sub imgColApp_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgApp.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgApp.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfApp, lngLeft, lngTop, imgColApp.Height)
    zl_vsGrid_Para_Save 1115, vsfApp, Me.Name, "vsfApp"
End Sub

Private Sub picImgApp_Click()
    Call imgColApp_Click
End Sub

Private Sub picApp_Resize()
    With vsfApp
        .Height = picApp.ScaleHeight
        .Width = picApp.ScaleWidth
    End With
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    With vsfList
        .Height = picList.ScaleHeight - 300
        .Width = picList.ScaleWidth
    End With
End Sub

Private Sub tabMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    cmdDel.Visible = False
    cmdChange.Visible = False
    If Item.index = 3 Then
        mstrPreNote = mfrmMain.stbThis.Panels(2).Text
        mfrmMain.stbThis.Panels(2).Text = ""
    End If
    If Item.index = 0 Then Call vsfList_EnterCell: mfrmMain.stbThis.Panels(2).Text = mstrPreNote
End Sub

Private Sub vsfList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfList
        If OldRow < .Rows Then
            If OldRow > 1 Then
                If OldRow Mod 2 = 0 Then
                    .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000005
                Else
                    .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &HEEEEEE
                End If
            End If
        End If
        If NewRow > 1 Then
            .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = 16772055
        End If
    End With
End Sub

Private Sub vsfApp_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfApp
        If OldRow < .Rows Then
            If OldRow Mod 2 = 1 Then
                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000005
            Else
                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &HEEEEEE
            End If
        End If
        .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = 16772055
    End With
End Sub

Private Sub vsfApp_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnMark Then mblnMark = False: Exit Sub
    With vsfApp
        Cancel = True
    End With
End Sub


Private Sub vsfList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call vsfList_EnterCell
End Sub

Private Sub vsfList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = 0 And Col = 0 Then
    Else
        Cancel = True
    End If
End Sub

Private Sub vsfList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsfList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = vsfList.ColIndex("预约时间") Then
        Call DeleteRecord
    Else
        Call SelectDoc
    End If
End Sub

Private Sub SelectDoc()
    Dim strArray() As String
    Dim lngRow As Long
    With vsfList
        lngRow = .Row
        frmServiceChangeNum.InitValue .TextMatrix(.Row, 1), .TextMatrix(.Row, 3), .TextMatrix(.Row, 4), _
            .TextMatrix(.Row, 5), .TextMatrix(.Row, 6), .TextMatrix(.Row, 9), lblInfo.Caption, _
            .Cell(flexcpData, .Row, 1)
        frmServiceChangeNum.mlng消息ID = Val(.RowData(.Row))
        frmServiceChangeNum.Show 1, Me
    End With

    Call LoadData(mlng消息ID)
    Call LocateNextRecord(lngRow)
End Sub

Private Sub vsfList_DblClick()
    With vsfList
        If .TextMatrix(.Row, .ColIndex("是否处理")) <> 0 Then Exit Sub
        If .Cell(flexcpForeColor, .Row, 1, .Row, 1) = vbRed Then Exit Sub
        If Val(.RowData(.Row)) = 0 Then Exit Sub
        If InStr(mstrPrivs, ";停诊信息处理;") = 0 Then Exit Sub
    End With
    Call mfrmMain.InformPatient
End Sub

Private Sub vsfList_EnterCell()
    Dim lngLeft As Long
    Dim i As Integer
    With vsfList
        If .Cell(flexcpForeColor, .Row, 1, .Row, 1) = vbRed Then
            cmdDel.Visible = False
            cmdChange.Visible = False
            Exit Sub
        End If
        If InStr(mstrPrivs, ";停诊信息处理;") = 0 Then
            cmdDel.Visible = False
            cmdChange.Visible = False
            Exit Sub
        End If
        If Val(.RowData(.Row)) = 0 Then Exit Sub
        cmdDel.Visible = True
        cmdChange.Visible = True
        cmdDel.Top = 318 * (.Row + 1)
        cmdChange.Top = cmdDel.Top
        For i = 0 To 9
            lngLeft = lngLeft + .ColWidth(i)
        Next i
        cmdDel.Left = lngLeft - cmdDel.Width
        cmdChange.Left = lngLeft + .ColWidth(10) - cmdChange.Width
    End With
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If vsfList.TextMatrix(vsfList.Row, 1) <> "" Then
            mcbrPopupMain.ShowPopup
        End If
    End If
End Sub

Private Sub vsfApp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        mcbrPopupMain.ShowPopup
    End If
End Sub
