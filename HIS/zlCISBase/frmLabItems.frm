VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmLabItems 
   Caption         =   "检验项目管理"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   Icon            =   "frmLabItems.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11205
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   6744
      Width           =   11208
      _ExtentX        =   19764
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabItems.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14684
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   90
      ScaleHeight     =   5295
      ScaleWidth      =   5310
      TabIndex        =   1
      Top             =   450
      Width           =   5310
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4410
         Left            =   0
         TabIndex        =   2
         Top             =   330
         Width           =   4395
         _Version        =   589884
         _ExtentX        =   7752
         _ExtentY        =   7779
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin VB.OptionButton optScope 
         Caption         =   "计算项目"
         Height          =   180
         Index           =   3
         Left            =   3840
         TabIndex        =   7
         Top             =   75
         Width           =   1065
      End
      Begin VB.OptionButton optScope 
         Caption         =   "组合项目"
         Height          =   180
         Index           =   2
         Left            =   2760
         TabIndex        =   6
         Top             =   75
         Width           =   1065
      End
      Begin VB.OptionButton optScope 
         Caption         =   "普通"
         Height          =   180
         Index           =   1
         Left            =   2055
         TabIndex        =   5
         Top             =   75
         Width           =   690
      End
      Begin VB.OptionButton optScope 
         Caption         =   "全部"
         Height          =   180
         Index           =   0
         Left            =   1350
         TabIndex        =   4
         Top             =   75
         Width           =   690
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   1845
         Top             =   4800
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
               Picture         =   "frmLabItems.frx":115C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLabItems.frx":16F6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblList 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "检验项目列表:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   3
         Top             =   90
         Width           =   1170
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1260
      Left            =   270
      TabIndex        =   8
      Top             =   5460
      Visible         =   0   'False
      Width           =   1305
      _cx             =   2302
      _cy             =   2222
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
      BackColorFixed  =   15790320
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
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
      AutoResize      =   0   'False
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabItems.frx":1C90
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    图标 = 0: ID: 类型: 编码: 中文名: 缩写: 标本: 单位: 组合: 性质: 结果类型
End Enum

Const conPane_List = 201
Const conPane_Base = 202
Const conPane_Ref = 203
Const conPane_Option = 204
Const conPane_Cost = 205
Const conPane_Merge = 206       ' 20070425
Const conpane_Significance = 207 '20081224  临床意义页
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串

Private mfrmBase As frmLabItemBase
Attribute mfrmBase.VB_VarHelpID = -1
Private mfrmRef As frmLabItemRef
Private mfrmSons As frmLabItemSons
Private mfrmOption As frmLabItemOption
Private mfrmCost As frmLabItemCost

Private mfrmMerge As frmLabItemMerge '20070425
Private mfrmSigni As frmLabItemSignificance  '2008-12-24

Public mblnShowStop As Boolean     '显示停用项目,设置为公共变量，以便查找窗体使用
Private mintEditState As Integer    '当前编辑状态：0-非编辑状态,1-基本信息编辑,2-参考值编辑,3-组成项编辑;4-执行选项编辑;5-试剂耗用;6-合并规则设置
Private mlngItemID As Long, mbln组合 As Boolean, mbln微生物 As Boolean
Private mLngEditWidth As Long       '为适应大字体情况下窗体变大.先读入窗体大小.
'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long



'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Public Function zlRefList(Optional lngItemID As Long) As Long
    '功能：刷新装入指定种类的病历文件清单，并定位到指定的文件上
    Dim rsTemp As New ADODB.Recordset
    Dim strGroups As String, blnShowIt As Boolean
    Me.rptList.Tag = ""
    gstrSql = "Select I.ID, Nvl(K.编码, 'N ') || '-' || I.操作类型 As 类型, I.编码, I.名称 As 中文名, L.缩写, I.标本部位 As 标本," & vbNewLine & _
            "       I.计算单位 As 单位, I.组合项目 As 组合, L.项目类别 As 性质, L.结果类型, I.撤档时间" & vbNewLine & _
            "From 诊疗项目目录 I, 检验报告项目 R, 检验项目 L, 诊疗检验类型 K" & vbNewLine & _
            "Where I.ID = R.诊疗项目id(+) And R.报告项目id = L.诊治项目id(+) And I.操作类型 = K.名称(+) and R.细菌id is null  And I.类别 = 'C' And" & vbNewLine & _
            "      I.组合项目 = 0" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select I.ID, Nvl(K.编码, 'N ') || '-' || I.操作类型 As 类型, I.编码, I.名称 As 中文名, '' As 缩写, I.标本部位 As 标本," & vbNewLine & _
            "       I.计算单位 As 单位, I.组合项目, Null+0 As 性质, Null+0 As 结果类型, I.撤档时间" & vbNewLine & _
            "From 诊疗项目目录 I, 诊疗检验类型 K" & vbNewLine & _
            "Where I.类别 = 'C' And I.操作类型 = K.名称(+) And I.组合项目 = 1"
    
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.rptList.Records.DeleteAll
    With rsTemp
        strGroups = ""
        Do While Not .EOF
            blnShowIt = Format(!撤档时间, "yyyy-mm-dd") = "3000-01-01" Or IsNull(!撤档时间) Or mblnShowStop = True
            If Me.optScope(1).Value And blnShowIt Then
                blnShowIt = (Val("" & !性质) = 1 Or Val("" & !性质) = 2) And (Val("" & !组合) = 0)
            ElseIf Me.optScope(2).Value And blnShowIt Then
                blnShowIt = (Val("" & !组合) <> 0)
            ElseIf Me.optScope(3).Value And blnShowIt Then
                blnShowIt = (Val("" & !性质) = 3)
            End If
            If blnShowIt Then
                If InStr(1, strGroups, !类型) = 0 Then strGroups = strGroups & "," & !类型
                Set rptRcd = Me.rptList.Records.Add()
                If Format("" & !撤档时间, "yyyy-mm-dd") = "3000-01-01" Or IsNull(!撤档时间) Then
                    Set rptItem = rptRcd.AddItem("0"): rptItem.Icon = 0
                Else
                    Set rptItem = rptRcd.AddItem("1"): rptItem.Icon = 1
                End If
                rptRcd.AddItem CStr(!ID)
                rptRcd.AddItem CStr("" & !类型)
                rptRcd.AddItem CStr(!编码)
                rptRcd.AddItem CStr(!中文名)
                rptRcd.AddItem CStr("" & !缩写)
                rptRcd.AddItem CStr("" & !标本)
                rptRcd.AddItem CStr("" & !单位)
                If Val("" & !组合) = 0 Then
                    rptRcd.AddItem ""
                Else
                    rptRcd.AddItem "√"
                End If
                Select Case Val("" & !性质)
                Case 3: rptRcd.AddItem "3-计算项"
                Case 2: rptRcd.AddItem "2-微生物"
                Case Else: rptRcd.AddItem "1-普通"
                End Select
                Select Case Val("" & !结果类型)
                Case 1: rptRcd.AddItem "1-定量"
                Case 2: rptRcd.AddItem "2-定性"
                Case 3: rptRcd.AddItem "3-半定量"
                Case Else: rptRcd.AddItem ""
                End Select
            End If
            .MoveNext
        Loop
        If strGroups <> "" Then strGroups = Mid(strGroups, 2)
    End With
    With Me.rptList
        If UBound(Split(strGroups, ",")) < 1 Then
            .GroupsOrder.DeleteAll
        ElseIf .GroupsOrder.Count = 0 Then
            .GroupsOrder.Add .Columns.Find(mCol.类型)
            .GroupsOrder(0).SortAscending = True
        End If
        .Populate
    End With
    
    Dim rptParent As ReportRow
    If lngItemID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = lngItemID Then
                    Set rptParent = rptRow.ParentRow
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow Then
                If Not (rptRow Is rptParent) Then rptRow.Expanded = False
            End If
        Next
        Set Me.rptList.FocusedRow = Me.rptList.FocusedRow
    Else
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow Then rptRow.Expanded = False
        Next
    End If
'    mlngItemID = 0
    If Me.rptList.Rows.Count > 0 And (Me.rptList.FocusedRow Is Nothing) Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    
    zlRefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "共有" & Me.rptList.Records.Count & "条项目"
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlControl.RPTCopyToVSF(Me.rptList, Me.vfgList) Is Nothing Then Exit Sub
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "检验项目清单"
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
End Sub

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRetuId As Long
    
    '------------------------------------
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_Save:
        Select Case mintEditState   '0-非编辑状态,1-基本信息编辑,2-参考值编辑,3-组成项编辑;4-执行选项编辑;5-试剂耗用;6-合并规则设置;7-临床意义设置
        Case 1
            lngRetuId = mfrmBase.zlEditSave()
            If lngRetuId <> 0 Then
                mlngItemID = lngRetuId: Call zlRefList(mlngItemID)
                mintEditState = 0: Me.picList.Enabled = True
            End If
        Case 2
            lngRetuId = mfrmRef.zlEditSave()
            If lngRetuId <> 0 Then Call zlRefList(mlngItemID): mintEditState = 0: Me.picList.Enabled = True
        Case 3
            lngRetuId = mfrmSons.zlEditSave()
            If lngRetuId <> 0 Then Call zlRefList(mlngItemID): mintEditState = 0: Me.picList.Enabled = True
        Case 4
            lngRetuId = mfrmOption.zlEditSave()
            If lngRetuId <> 0 Then mintEditState = 0: Me.picList.Enabled = True
        Case 5
            lngRetuId = mfrmCost.zlEditSave()
            If lngRetuId <> 0 Then mintEditState = 0: Me.picList.Enabled = True
        Case 6
            lngRetuId = mfrmMerge.zlEditSave()
            If lngRetuId <> 0 Then mintEditState = 0: Me.picList.Enabled = True
        Case 7
            lngRetuId = mfrmSigni.zlEditSave
            If lngRetuId <> 0 Then mintEditState = 0: Me.picList.Enabled = True
        End Select
        
    Case conMenu_Edit_Untread:
        Select Case mintEditState   '0-非编辑状态,1-基本信息编辑,2-参考值编辑,3-组成项编辑;4-执行选项编辑;5-试剂耗用;6-合并规则设置
        Case 1: Call mfrmBase.zlEditCancel
        Case 2: Call mfrmRef.zlEditCancel
        Case 3: Call mfrmSons.zlEditCancel
        Case 4: Call mfrmOption.zlEditCancel
        Case 5: Call mfrmCost.zlEditCancel
        Case 6: Call mfrmMerge.zlEditCancel
        Case 7: Call mfrmSigni.zlEditCancel
        End Select
        mintEditState = 0: Me.picList.Enabled = True
    
    Case conMenu_Edit_NewItem
        If mfrmBase.zlEditStart(True, mlngItemID) Then mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Base).Select
    Case conMenu_Edit_Modify
        If mlngItemID = 0 Then Exit Sub
        If mfrmBase.zlEditStart(False, mlngItemID) Then mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Base).Select
    Case conMenu_Edit_Delete
        If mlngItemID = 0 Then Exit Sub
        
        ' 检查是否被其他项目引用
        Dim rsGS As ADODB.Recordset
        Dim strTmp As String, strItem As String
        
        gstrSql = "Select 诊治项目id, 缩写, B.中文名, B.编码" & vbNewLine & _
                "From 诊治所见项目 B, 检验项目 A" & vbNewLine & _
                "Where A.诊治项目id = B.ID And" & vbNewLine & _
                "      计算公式 Like (Select '%' || Chr(91) || A.报告项目id || Chr(93) || '%' From 检验报告项目 A ,诊疗项目目录 B Where A.诊疗项目id=B.ID and B.组合项目=0 and A.诊疗项目id = [1])"
        Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID)
        strTmp = "此项目被以下项目引用，不能删除！"
        Do Until rsGS.EOF
            strItem = strItem & "(" & rsGS.Fields("编码") & ")" & rsGS.Fields("中文名") & vbNewLine
            rsGS.MoveNext
        Loop
        If strItem <> "" Then
            MsgBox strTmp & vbNewLine & strItem, vbInformation, Me.Caption
            Exit Sub
        End If
        
        With Me.rptList
            If MsgBox("真的删除该检验项目吗？" & vbCrLf & "——" & .FocusedRow.Record(mCol.中文名).Value, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSql = "Zl_检验项目_Edit(3," & mlngItemID & ")"
                Err = 0: On Error GoTo ErrHand
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
                
                Err = 0: On Error GoTo 0
                mlngItemID = 0: lngRetuId = .FocusedRow.Index
                If .Rows.Count > lngRetuId + 1 Then
                    If .Rows(lngRetuId + 1).GroupRow = False Then mlngItemID = .Rows(lngRetuId + 1).Record(mCol.ID).Value
                ElseIf lngRetuId > 0 Then
                    If .Rows(lngRetuId - 1).GroupRow = False Then mlngItemID = .Rows(lngRetuId - 1).Record(mCol.ID).Value
                End If
                Call Me.zlRefList(mlngItemID)
            End If
        End With
        Exit Sub
    
    Case conMenu_Edit_Compend
        If mbln组合 Or mbln微生物 Then
            If mfrmSons.zlEditStart Then mintEditState = 3: Me.picList.Enabled = False
        Else
            If mfrmRef.zlEditStart Then mintEditState = 2: Me.picList.Enabled = False
        End If
        Me.dkpMan.FindPane(conPane_Ref).Select
    Case conMenu_Edit_ApplyTo
        If mfrmOption.zlEditStart Then mintEditState = 4: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Option).Select
    Case conMenu_Edit_Test
        If mfrmCost.zlEditStart Then mintEditState = 5: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Cost).Select
    Case conMenu_Edit_Merge '20070427
        If mfrmBase.chk单独应用 = 1 Then
            If mfrmMerge.zlEditStart Then mintEditState = 6: Me.picList.Enabled = False
            Me.dkpMan.FindPane(conPane_Merge).Select
        End If
    Case conMenu_Edit_Sort '20080722 项目排序
        frmItemSort.Show vbModal, Me
    Case conMenu_Edit_Import '2008-12-24 临床意义

        If mfrmSigni.zlEditStart Then mintEditState = 7: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conpane_Significance).Select
    Case conMenu_Edit_Pause
        With Me.rptList
            If MsgBox("真的要停用该检验项目吗？" & vbCrLf & "——" & .FocusedRow.Record(mCol.中文名).Value, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSql = "zl_诊疗项目_STOP(" & mlngItemID & ")"
                Err = 0: On Error GoTo ErrHand
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            End If
        End With
        Call Me.zlRefList(mlngItemID)
    Case conMenu_Edit_Reuse
        With Me.rptList
            If MsgBox("真的重新启用该检验项目吗？" & vbCrLf & "——" & .FocusedRow.Record(mCol.中文名).Value, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSql = "zl_诊疗项目_REUSE(" & mlngItemID & ")"
                Err = 0: On Error GoTo ErrHand
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            End If
        End With
        Call Me.zlRefList(mlngItemID)
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Find
        frmLabItemFind.Show , Me
    Case conMenu_View_Refresh
        Call zlRefList(mlngItemID)
    Case conMenu_View_Option
        mblnShowStop = Not mblnShowStop: Control.Checked = mblnShowStop
        Call zlRefList(mlngItemID)
    
    Case conMenu_Help_Help:     Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Dim lngItemID As Long
    If Me.rptList.FocusedRow Is Nothing Then
        lngItemID = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        lngItemID = 0
    Else
        lngItemID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel: Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread: Control.Enabled = (mintEditState <> 0)
    
    Case conMenu_Edit_NewItem: Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Compend, conMenu_Edit_ApplyTo, conMenu_Edit_Sort
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0 And lngItemID <> 0)
    Case conMenu_Edit_Import
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0 And lngItemID <> 0) And Not mbln组合
    Case conMenu_Edit_Test
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0 And lngItemID <> 0)
        If Control.Enabled Then Control.Enabled = Not mbln组合
    Case conMenu_Edit_Pause
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0 And lngItemID <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record.Item(mCol.图标).Value = 0)
    Case conMenu_Edit_Reuse
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0 And lngItemID <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record.Item(mCol.图标).Value <> 0)
    
    Case conMenu_Edit_Merge '20070425
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0 And lngItemID <> 0)
        If Control.Enabled And (Not mfrmBase Is Nothing) Then Control.Enabled = mfrmBase.chk单独应用
        
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Base
        If mfrmBase Is Nothing Then Set mfrmBase = New frmLabItemBase
        Item.Handle = mfrmBase.hWnd
    Case conPane_Ref
        If mbln组合 Or mbln微生物 Then
            If mfrmSons Is Nothing Then Set mfrmSons = New frmLabItemSons
            Item.Handle = mfrmSons.hWnd
        Else
            If mfrmRef Is Nothing Then Set mfrmRef = New frmLabItemRef
            Item.Handle = mfrmRef.hWnd
        End If
    Case conPane_Option
        If mfrmOption Is Nothing Then Set mfrmOption = New frmLabItemOption
        Item.Handle = mfrmOption.hWnd
    Case conPane_Cost
        If mfrmCost Is Nothing Then Set mfrmCost = New frmLabItemCost
        Item.Handle = mfrmCost.hWnd
    Case conPane_Merge '20070425
        If mfrmMerge Is Nothing Then Set mfrmMerge = New frmLabItemMerge
        Item.Handle = mfrmMerge.hWnd
        
        If Not mfrmBase Is Nothing Then
            If mfrmBase.chk单独应用 = 1 Then
                Item.Hidden = False
            Else
                Item.Hidden = True
            End If
        End If
    Case conpane_Significance
        If mfrmSigni Is Nothing Then Set mfrmSigni = New frmLabItemSignificance
        Item.Handle = mfrmSigni.hWnd
    End Select
End Sub

Private Sub Form_Load()
    
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    
    mLngEditWidth = frmLabItemBase.ScaleWidth
    
    mintEditState = 0: mblnShowStop = False
    mlngItemID = 0: mbln组合 = False: mbln微生物 = False
    For lngCount = Me.optScope.LBound To Me.optScope.UBound
        Me.optScope(lngCount).BackColor = Me.picList.BackColor
    Next
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    'Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "参考或组成(&B)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "执行选项(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Test, "试剂比例(&G)")
        
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Merge, "合并规则(&M)") ' --  20070425
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Sort, "项目顺序(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "临床意义(&L)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Pause, "暂停(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用(&U)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Find, "查找(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "显示停用(&H)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("B"), conMenu_Edit_Compend
        .Add FCONTROL, Asc("E"), conMenu_Edit_ApplyTo
        .Add FCONTROL, Asc("G"), conMenu_Edit_Test
        .Add FCONTROL, Asc("F"), conMenu_View_Find
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_Edit_Pause
        .AddHiddenCommand conMenu_Edit_Reuse
        .AddHiddenCommand conMenu_View_Refresh
        .AddHiddenCommand conMenu_View_Option
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "参考或组成"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "执行选项")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Test, "试剂比例")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Merge, "合并规则") '--20070425
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Sort, "项目顺序")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "临床意义")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panThis As Pane, panSub1 As Pane, panSub2 As Pane
    
    If mfrmBase Is Nothing Then Set mfrmBase = New frmLabItemBase
    If mfrmRef Is Nothing Then Set mfrmRef = New frmLabItemRef
    If mfrmSons Is Nothing Then Set mfrmSons = New frmLabItemSons
    If mfrmOption Is Nothing Then Set mfrmOption = New frmLabItemOption
    If mfrmCost Is Nothing Then Set mfrmCost = New frmLabItemCost
    If mfrmMerge Is Nothing Then Set mfrmMerge = New frmLabItemMerge
    If mfrmSigni Is Nothing Then Set mfrmSigni = New frmLabItemSignificance
    
    Set panThis = dkpMan.CreatePane(conPane_List, 450, 580, DockLeftOf, Nothing)
    panThis.Title = "检验项目列表"
    panThis.Options = PaneNoCaption
    
    Set panThis = dkpMan.CreatePane(conPane_Base, 550, 580, DockRightOf, Nothing)
    panThis.Title = "项目基本属性"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panSub1 = dkpMan.CreatePane(conPane_Ref, 550, 800, DockBottomOf, panThis)
    panSub1.Title = "项目参考值"
    panSub1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panSub2 = dkpMan.CreatePane(conPane_Option, 550, 800, DockBottomOf, panSub1)
    panSub2.Title = "项目执行选项"
    panSub2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panSub2.AttachTo panSub1
    
    Set panSub2 = dkpMan.CreatePane(conPane_Cost, 550, 800, DockBottomOf, panSub1)
    panSub2.Title = "项目试剂比例"
    panSub2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panSub2.AttachTo panSub1
    
    Set panSub2 = dkpMan.CreatePane(conPane_Merge, 550, 800, DockBottomOf, panSub1)
    panSub2.Title = "合并规则设置"
    panSub2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panSub2.AttachTo panSub1
    
    Set panSub2 = dkpMan.CreatePane(conpane_Significance, 550, 800, DockBottomOf, panSub1)
    panSub2.Title = "临床意义"
    panSub2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panSub2.AttachTo panSub1
    
    panSub1.Select
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 1024)   '必须在列设置之前设置，才能生效
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.类型, "类型", 70, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编码, "编码", 60, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.中文名, "中文名", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.缩写, "缩写", 70, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.标本, "标本", 70, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.单位, "单位", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.组合, "组合", 30, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.性质, "性质", 50, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.结果类型, "结果类型", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        
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
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '数据装入
    Dim intScope As Integer     '列表显示范围
    
    intScope = Abs(zlDatabase.GetPara("列表范围", glngSys, 1059, 0))
    If intScope < 4 Then
        Me.optScope(intScope).Value = True
    Else
        Me.optScope(0).Value = True
    End If
'    Call zlRefList

End Sub

Private Sub Form_Resize()
    Dim panBase As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panBase = Me.dkpMan.FindPane(conPane_Base)
    panBase.MinTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 265
    panBase.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 375
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters

    panBase.MinTrackSize.SetSize 0, 0
    panBase.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 375
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intScope As Integer
    If Me.optScope(0).Value Then
        intScope = 0
    ElseIf Me.optScope(1).Value Then
        intScope = 1
    ElseIf Me.optScope(1).Value Then
        intScope = 2
    ElseIf Me.optScope(1).Value Then
        intScope = 3
    Else
        intScope = 0
    End If
    Call zlDatabase.SetPara("列表范围", intScope, glngSys, 1059)
    Unload mfrmBase
    Unload mfrmRef
    Unload mfrmSons
    Unload mfrmOption
    Unload mfrmCost
    Set mfrmBase = Nothing
    Set mfrmRef = Nothing
    Set mfrmSons = Nothing
    Set mfrmOption = Nothing
    Set mfrmCost = Nothing
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub optScope_Click(Index As Integer)
    Dim lngItemID As Long
    
    If Me.rptList.FocusedRow Is Nothing Then
        lngItemID = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        lngItemID = 0
    Else
        lngItemID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    
    Call Me.zlRefList(lngItemID)
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Height = Me.picList.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(mCol.ID))
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim lngItemID As Long
    If Me.rptList.FocusedRow Is Nothing Then
        lngItemID = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        lngItemID = 0
    Else
        lngItemID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    If lngItemID = 0 Then Exit Sub
    
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)

End Sub

Private Sub rptList_SelectionChanged()
    If Me.rptList.FocusedRow Is Nothing Then
        mlngItemID = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngItemID = 0
    Else
        mlngItemID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        mbln组合 = (Me.rptList.FocusedRow.Record.Item(mCol.组合).Value = "√")
        
        mbln微生物 = (Me.rptList.FocusedRow.Record.Item(mCol.性质).Value = "2-微生物")
    End If
    
    Dim panThis As Pane, panSub1 As Pane
    
    Set panThis = Me.dkpMan.FindPane(conPane_Ref)
    If mbln组合 Or mbln微生物 Then
        If panThis.Handle <> mfrmSons.hWnd Then
            panThis.Handle = mfrmSons.hWnd
            mfrmRef.Visible = False
            panThis.Title = "组成项目列表"
            
            Set panSub1 = Me.dkpMan.FindPane(conPane_Option)
            panSub1.AttachTo panThis
            
            Set panSub1 = Me.dkpMan.FindPane(conPane_Cost)
            panSub1.AttachTo panThis
            
            
            Set panSub1 = Me.dkpMan.FindPane(conPane_Merge)
            If dkpMan.FindPane(conPane_Merge).Closed Then dkpMan.ShowPane conPane_Merge
            panSub1.AttachTo panThis
            
            Set panSub1 = Me.dkpMan.FindPane(conpane_Significance)
            If dkpMan.FindPane(conpane_Significance).Closed Then dkpMan.ShowPane conpane_Significance
            panSub1.AttachTo panThis
            
            panThis.Select
            Me.dkpMan.RecalcLayout
            
        End If
    Else
        If panThis.Handle <> mfrmRef.hWnd Then
            mfrmRef.Visible = True
            panThis.Handle = mfrmRef.hWnd
            panThis.Title = "项目参考值"
            
            Set panSub1 = Me.dkpMan.FindPane(conPane_Option)
            panSub1.AttachTo panThis
            
            Set panSub1 = Me.dkpMan.FindPane(conPane_Cost)
            panSub1.AttachTo panThis
            
            Set panSub1 = Me.dkpMan.FindPane(conPane_Merge)
            If dkpMan.FindPane(conPane_Merge).Closed Then dkpMan.ShowPane conPane_Merge
            panSub1.AttachTo panThis
            
            Set panSub1 = Me.dkpMan.FindPane(conpane_Significance)
            If dkpMan.FindPane(conpane_Significance).Closed Then dkpMan.ShowPane conpane_Significance
            panSub1.AttachTo panThis
            
            panThis.Select
            Me.dkpMan.RecalcLayout
            
        End If
    End If
    
    Call mfrmBase.zlRefresh(mlngItemID)
    If mbln组合 Or mbln微生物 Then
        Call mfrmSons.zlRefresh(mlngItemID)
    Else
        Call mfrmRef.zlRefresh(mlngItemID)
    End If
    Call mfrmOption.zlRefresh(mlngItemID)
    If mbln组合 Then
        Call mfrmCost.zlRefresh(0)
        If Not dkpMan.FindPane(conpane_Significance).Closed Then
            dkpMan.FindPane(conpane_Significance).Close
            panThis.Select
            Me.dkpMan.RecalcLayout
        End If
        Call mfrmSigni.zlRefresh(0)
    Else
        Call mfrmCost.zlRefresh(mlngItemID)
        If dkpMan.FindPane(conpane_Significance).Closed Then
            dkpMan.ShowPane conpane_Significance
            panThis.Select
            Me.dkpMan.RecalcLayout
        End If
        Call mfrmSigni.zlRefresh(mlngItemID)
    End If
    
    
    If mfrmBase.chk单独应用 = 0 Then
         If Not dkpMan.FindPane(conPane_Merge).Closed Then
            dkpMan.FindPane(conPane_Merge).Close
            Me.dkpMan.RecalcLayout
        End If
        Call mfrmMerge.zlRefresh(0)
    Else
        If dkpMan.FindPane(conPane_Merge).Closed Then
            dkpMan.ShowPane conPane_Merge
            panThis.Select
            Me.dkpMan.RecalcLayout
        End If
        
        Call mfrmMerge.zlRefresh(mlngItemID)
    End If

    
End Sub
