VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmEPRModelList 
   BorderStyle     =   0  'None
   Caption         =   "病历范文"
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3090
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4905
      _Version        =   589884
      _ExtentX        =   8652
      _ExtentY        =   5450
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   165
      Top             =   3045
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelList.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelList.frx":0B34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3765
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
End
Attribute VB_Name = "frmEPRModelList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const con_UnDefine = -999

Const conColumn_图标 = 0
Const conColumn_ID = 1
Const conColumn_编号 = 2
Const conColumn_名称 = 3
Const conColumn_说明 = 4
Const conColumn_部门 = 5
Const conColumn_人员 = 6

'---------------------------------
'公共事件
'---------------------------------
Public Event RightMouseUp(X As Long, Y As Long)                         '右键鼠标事件
Public Event RowDblClick(ByVal Row As XtremeReportControl.IReportRow)   '双击一行或在行上按回车
Public Event SelRowChanged(ByVal Row As XtremeReportControl.IReportRow) '选择行改变

'---------------------------------
'窗体变量
'---------------------------------
Private mlngFileID As Long           '当前指定的文件id
Private mintPower As Integer        '词句管理权范围
'    mintPower=con_UnDefine，未定义;
'    mintPower=-1，不具备词句管理权;
'    mintPower=0，全院，这时显示所有的示范，也可以更改;
'    mintPower=1，科室，这时显示全院通用示范(科室id is null)和所在科室公有或部门内人员私有的示范，但不能更改全院通用示范;
'    mintPower=2，个人，这时显示全院通用示范(科室id is null)和所在科室通用示范(人员id is null)和个人示范，仅个人示范可更改

'---------------------------------
'临时变量
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow


Private Sub Form_Activate()
    Err = 0: On Error Resume Next
    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Me.rptList.SetFocus
End Sub

Private Sub Form_Load()
    mintPower = con_UnDefine
    Call zlGetPower
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(conColumn_图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(conColumn_ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(conColumn_编号, "编号", 40, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(conColumn_名称, "名称", 110, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(conColumn_说明, "说明", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(conColumn_部门, "部门", 70, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(conColumn_人员, "编制人", 50, False): rptCol.Editable = False: rptCol.Groupable = True
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
End Sub

Private Sub Form_Resize()
    With Me.rptList
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(conColumn_编号))
End Sub

Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call Form_Activate
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button <> vbRightButton Then Exit Sub
    RaiseEvent RightMouseUp(X, Y)
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    RaiseEvent RowDblClick(Row)
End Sub

Private Sub rptList_SelectionChanged()
    If Me.rptList.Visible = False Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.Tag = "" And (Val(Me.rptList.Tag) <> Me.rptList.FocusedRow.Index Or Me.rptList.Tag = "") Then
        RaiseEvent SelRowChanged(Me.rptList.FocusedRow)
        Me.rptList.Tag = Me.rptList.FocusedRow.Index
    End If
End Sub

'-----------------------------------------------------
'窗体公共方法
'-----------------------------------------------------

Public Function zlRefresh(ByVal lngFileID As Long, Optional ByVal lngRecId As Long) As Long
    '功能：刷新装入指定文件的范文目录
    '参数： lngFileId，文件ID
    '       lngRecId，需要定位到的范文
    '返回：刷新装入的范文数目
    Me.Tag = "zlRefresh"
    Me.rptList.Tag = ""
    mlngFileID = lngFileID
    Err = 0: On Error GoTo ErrHand
    Select Case mintPower
    Case 0
        gstrSQL = "Select l.Id, l.编号, l.名称, l.说明, l.通用级, d.名称 As 部门, p.姓名 As 人员 " _
                & "From 病历范文目录 l, 部门表 d, 人员表 p " _
                & "Where l.科室id = d.Id And l.人员id = p.Id And l.文件id =[1] " _
                & "Order By l.编号"
    Case 1
        gstrSQL = "Select l.Id, l.编号, l.名称, l.说明, l.通用级, d.名称 As 部门, p.姓名 As 人员 " _
                & "From 病历范文目录 l, 部门表 d, 人员表 p " _
                & "Where l.科室id = d.Id(+) And l.人员id = p.Id(+) And l.文件id =[1] And " _
                & "      (Nvl(l.通用级, 0) = 0 Or " _
                & "      l.通用级 in (1,2) And l.科室id In (Select r.部门id From 部门人员 r, 上机人员表 u Where r.人员id = u.人员id And u.用户名 = User)) " _
                & "Order By l.编号"
    Case Else
        gstrSQL = "Select l.Id, l.编号, l.名称, l.说明, l.通用级, d.名称 As 部门, p.姓名 As 人员 " _
                & "From 病历范文目录 l, 部门表 d, 人员表 p " _
                & "Where l.科室id = d.Id(+) And l.人员id = p.Id(+) And l.文件id =[1]  And " _
                & "      (Nvl(l.通用级, 0) = 0 Or " _
                & "      l.通用级 =1 And l.科室id In (Select r.部门id From 部门人员 r, 上机人员表 u Where r.人员id = u.人员id And u.用户名 = User) Or " _
                & "      l.通用级 =2 And l.人员id In (Select u.人员id From 上机人员表 u Where u.用户名 = User)) " _
                & "Order By l.编号"
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    
    Me.rptList.Records.DeleteAll
    Do While Not rsTemp.EOF
        Set rptRcd = Me.rptList.Records.Add()
        Set rptItem = rptRcd.AddItem(CStr(IIf(IsNull(rsTemp!通用级), 0, rsTemp!通用级)))
        rptItem.Icon = rptItem.Value: rptItem.GroupPriority = rptItem.Value: rptItem.SortPriority = rptItem.Value
        Select Case rptItem.Value
        Case 0: rptItem.GroupCaption = "1-全院"
        Case 1: rptItem.GroupCaption = "2-科室"
        Case Else: rptItem.GroupCaption = "3-个人"
        End Select
        rptRcd.AddItem CStr(rsTemp!ID)
        rptRcd.AddItem Val(CStr(rsTemp!编号))
        rptRcd.AddItem CStr(rsTemp!名称)
        rptRcd.AddItem CStr("" & rsTemp!说明)
        rptRcd.AddItem CStr("" & rsTemp!部门)
        rptRcd.AddItem CStr("" & rsTemp!人员)
        rsTemp.MoveNext
    Loop
    Me.rptList.Populate
    
    If Me.rptList.Rows.Count > 0 Then
        If lngRecId <> 0 Then
            For Each rptRow In Me.rptList.Rows
                Set Me.rptList.FocusedRow = rptRow
            Next
        End If
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Me.Tag = ""
    Call rptList_SelectionChanged
    zlRefresh = Me.rptList.Records.Count
    
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = Me.rptList.Records.Count
End Function

Public Function zlGetPower() As Integer
    '功能：获得当前用户的范文管理的权限
    '返回：词句管理权限数值
    Dim strPrivs As String
    If mintPower = con_UnDefine Then
        strPrivs = GetPrivFunc(glngSys, 1070)
        If InStr(1, strPrivs, "全院病历范文") <> 0 Then
            mintPower = 0
        ElseIf InStr(1, strPrivs, "科室病历范文") <> 0 Then
            mintPower = 1
        ElseIf InStr(1, strPrivs, "个人病历范文") <> 0 Then
            mintPower = 2
        Else
            mintPower = -1
        End If
    End If
    zlGetPower = mintPower
End Function

Public Function zlGetFocusedRow() As XtremeReportControl.IReportRow
    '功能：获取当前选中行
    If Me.rptList.FocusedRow Is Nothing Then
        Set zlGetFocusedRow = Nothing
    Else
        Set zlGetFocusedRow = Me.rptList.FocusedRow
    End If
End Function

Public Function zlRecordCount() As Long
    '功能:返回当前的记录总数
    zlRecordCount = Me.rptList.Records.Count
End Function

Public Function zlRecordNew(ByVal frmParent As Object) As Long
    '功能：增加新的范文
    Dim lngItemId As Long
    If frmParent Is Nothing Then Set frmParent = Me
    lngItemId = frmEPRModelEdit.ShowMe(frmParent, True, CByte(mintPower), mlngFileID)
    If lngItemId <> 0 Then Call Me.zlRefresh(mlngFileID, lngItemId)
    zlRecordNew = lngItemId
    DoEvents: Call Form_Activate
End Function

Public Function zlRecordModify(ByVal frmParent As Object) As Long
    '功能：当前范文修改
    Dim lngItemId As Long
    If Me.rptList.FocusedRow Is Nothing Then zlRecordModify = 0: Exit Function
    If Me.rptList.FocusedRow.GroupRow = True Then zlRecordModify = 0: Exit Function
    lngItemId = Me.rptList.FocusedRow.Record.Item(conColumn_ID).Value
    If frmParent Is Nothing Then Set frmParent = Me
    
    lngItemId = frmEPRModelEdit.ShowMe(frmParent, False, CByte(mintPower), mlngFileID, lngItemId)
    If lngItemId <> 0 Then Call Me.zlRefresh(mlngFileID, lngItemId)
    zlRecordModify = lngItemId
    DoEvents: Call Form_Activate
End Function

Public Function zlRecordDelete() As Boolean
    '功能：当前范文删除
    Dim lngIndex As Long, lngItemId As Long
    
    With Me.rptList
        If .FocusedRow Is Nothing Then Exit Function
        If .FocusedRow.GroupRow Then Exit Function
        
        If MsgBox("真的删除该范文吗？" & vbCrLf & "——" & .FocusedRow.Record(conColumn_名称).Value, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        gstrSQL = "zl_病历范文目录_delete('" & .FocusedRow.Record(conColumn_ID).Value & "')"
        Err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Err = 0: On Error GoTo 0
        lngIndex = .FocusedRow.Record.Index
        Call .Records.RemoveAt(.FocusedRow.Record.Index)
        .Populate
        If .Records.Count <> 0 Then
            If lngIndex >= .Records.Count Then lngIndex = 0
            lngItemId = .Records(lngIndex).Item(conColumn_ID).Value
        Else
            lngItemId = 0
        End If
        Call Me.zlRefresh(mlngFileID, lngItemId)
    End With
    zlRecordDelete = True
    DoEvents: Call Form_Activate
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL

    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vgdList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String

    Err = 0: On Error GoTo ErrHand
    gstrSQL = "Select 名称 From 病历文件列表 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If rsTemp.RecordCount > 0 Then strSubhead = rsTemp!名称
    
    Err = 0: On Error Resume Next
    Set objPrint.Body = Me.vgdList
    objPrint.Title.Text = strSubhead & "范文目录"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
