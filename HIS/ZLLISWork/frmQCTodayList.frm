VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmQCTodayList 
   Caption         =   "今日质控管理"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   Icon            =   "frmQCTodayList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10470
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picLeft 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   6750
      Left            =   90
      ScaleHeight     =   6750
      ScaleWidth      =   6060
      TabIndex        =   3
      Top             =   570
      Width           =   6060
      Begin VB.Frame fraNS 
         BorderStyle     =   0  'None
         Height          =   45
         Left            =   -225
         MousePointer    =   7  'Size N S
         TabIndex        =   7
         Top             =   2430
         Width           =   3360
      End
      Begin VB.PictureBox PicList 
         BorderStyle     =   0  'None
         Height          =   4200
         Left            =   345
         ScaleHeight     =   4200
         ScaleWidth      =   5610
         TabIndex        =   4
         Top             =   120
         Width           =   5610
         Begin XtremeReportControl.ReportControl rptList 
            Height          =   3210
            Left            =   90
            TabIndex        =   5
            Top             =   90
            Width           =   5280
            _Version        =   589884
            _ExtentX        =   9313
            _ExtentY        =   5662
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgList 
            Height          =   900
            Left            =   0
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   3300
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
         Begin MSComctlLib.ImageList imgList 
            Left            =   1245
            Top             =   3510
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
                  Picture         =   "frmQCTodayList.frx":058A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmQCTodayList.frx":0B24
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgRecord 
         Height          =   2295
         Left            =   720
         TabIndex        =   8
         Top             =   4605
         Width           =   4305
         _cx             =   7594
         _cy             =   4048
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
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
   Begin VB.ComboBox cbo仪器 
      Height          =   300
      Left            =   5550
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   2115
   End
   Begin VB.ComboBox cbo科室 
      Height          =   300
      Left            =   3660
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   1845
   End
   Begin MSComCtl2.DTPicker dtp日期 
      Height          =   300
      Left            =   7860
      TabIndex        =   0
      Top             =   15
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   98566147
      CurrentDate     =   39110
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7380
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQCTodayList.frx":10BE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13388
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   90
      Top             =   15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmQCTodayList.frx":1950
      Left            =   840
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmQCTodayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    图标 = 0: 标本ID: 标本号:  仪器id: 检验仪器: 质控品id: 质控品: 批号: 水平: 次数
End Enum
Private Enum mColL
    图标 = 0: ID: 中文名: 英文名: 结果: 靶值: SD: 单位: 序号: 取值序列: 弃用结果: 项目id: 质控品id: 开始日期: 结束日期: 原始结果: 归档人: 标记
End Enum
Const conPane_List = 201
Const conPane_LJ = 202
Const conPane_Report = 203

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串

Private mfrmLJ As frmQCChartLJ
Private mfrmReport As frmQCTodayReport

Private mintEditState As Integer    '当前编辑状态：0-非编辑状态,1-质控记录编辑,2-报告编辑
Private mstrDate As String          '日期
Private mlngRecord As Long          '样本id
Private mlngResult As Long          '结果id

Private mblnAllDev As Boolean      '是否具备所有仪器权限，否则只能处理本部门的仪器
Private mlngEditWidth As Long, mlngEditHeight As Long   '编辑窗格的高度和宽度

'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrCustom As CommandBarControlCustom
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim RptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long
Dim mblnEdit As Boolean '是否编辑过

'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Public Function zlRefList(Optional lngRecord As Long, Optional lngResult As Long) As Long
    '功能：刷新装入指定种类的病历文件清单，并定位到指定的记录
    Dim rsTemp As New ADODB.Recordset
    Dim strLists As String, strValue As String
    
    mstrDate = Format(Me.dtp日期.Value, "yyyy-MM-dd")
    If Me.cbo科室.Tag <> "" Then Exit Function
    Err = 0: On Error GoTo ErrHand
                
   gstrSql = " Select l.标本id, l.标本序号 as 标本号, l.仪器id,m.名称 as 仪器, x.名称 as 质控品, l.质控品id, x.批号, x.水平, l.测试次数 as 次数 " & vbNewLine & _
            " From 检验质控记录 l, 检验质控品 x,检验仪器 m " & vbNewLine & _
            " Where (l.检验时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([1], 'yyyy-mm-dd') + 1 - 1 / 86400) And" & vbNewLine & _
            " l.仪器id = m.id And l.质控品id = X.ID "
    
    '仪器
    If Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex) > 0 Then
        gstrSql = gstrSql & " and  L.仪器id = [2] "
    End If
    gstrSql = gstrSql & " Order by L.标本序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrDate, Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex))
    
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
'          Select Case Val("" & !标记)
'            Case 1
'                Set RptItem = rptRcd.AddItem("1"): RptItem.Icon = 0
'            Case 2
'                Set RptItem = rptRcd.AddItem("2"): RptItem.Icon = 1
'            Case Else
                Set RptItem = rptRcd.AddItem("")
'            End Select
            rptRcd.AddItem CLng(!标本ID)
            rptRcd.AddItem CStr("" & !标本号)
            rptRcd.AddItem CStr("" & !仪器id)
            rptRcd.AddItem CStr("" & !仪器)
            rptRcd.AddItem CStr("" & !质控品id)
            rptRcd.AddItem CStr("" & !质控品)
            rptRcd.AddItem CStr("" & !批号)
            rptRcd.AddItem CStr("" & !水平)
            rptRcd.AddItem CStr("" & !次数)
            .MoveNext
        Loop
    End With
    Me.rptList.Populate
    
    If lngRecord <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.标本ID).Value) = lngRecord And lngRecord <> 0 Then
                    Set Me.rptList.FocusedRow = rptRow: If mlngResult = 0 Then Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.Rows.Count > 0 And (Me.rptList.FocusedRow Is Nothing) Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Call rptList_SelectionChanged
    zlRefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "共有" & Me.rptList.Records.Count & "条记录"
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
    'If zlReportToVSFlexGrid(Me.vfgList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = Format(mstrDate, "yyyy年MM月dd日") & "质控记录清单"
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

Private Sub cbo科室_Click()
    Dim rsTmp As New ADODB.Recordset
    
'    gstrSql = "Select id ,编码 , 名称 From 检验仪器 a Where 使用小组id = [1] order by 编码 "
    gstrSql = "Select ID, 编码, 名称" & vbNewLine & _
            " From 检验仪器 A" & vbNewLine & _
            " Where 使用小组id = [1] And" & vbNewLine & _
            "      A.ID In (Select Distinct D.ID" & vbNewLine & _
            "               From 检验小组成员 A, 检验小组 B, 检验小组仪器 C, 检验仪器 D" & vbNewLine & _
            "               Where A.小组id = B.ID And B.ID = C.小组id　and 人员id = [2] And C.仪器id = D.ID)" & vbNewLine & _
            " Order By 编码"


    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, Me.cbo科室.ItemData(Me.cbo科室.ListIndex), UserInfo.ID)
    Me.cbo仪器.Clear
    If InStr(1, mstrPrivs, "所有科室") > 0 Then
        Me.cbo仪器.AddItem "所有仪器"
        Me.cbo仪器.ItemData(Me.cbo仪器.NewIndex) = 0
    End If
    
    Do Until rsTmp.EOF
        Me.cbo仪器.AddItem rsTmp("名称")
        Me.cbo仪器.ItemData(Me.cbo仪器.NewIndex) = rsTmp("ID")
        rsTmp.MoveNext
    Loop
    Me.cbo仪器.ListIndex = 0

    
End Sub

Private Sub cbo仪器_Click()
    Call zlRefList
End Sub

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRetuId As Long, strInfo As String
    
    '------------------------------------
    Select Case Control.ID
'    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_Save
        lngRetuId = 0
        Select Case mintEditState   '当前编辑状态：0-非编辑状态,1-质控记录编辑,2-报告编辑
        Case 1
            lngRetuId = zlEditSave()
            If lngRetuId <> 0 Then
                mlngRecord = lngRetuId ':  Call zlRefList(mlngRecord)
                
                mintEditState = 0: Me.dtp日期.Enabled = True
                Me.cbo科室.Enabled = True: Me.cbo仪器.Enabled = True: Me.PicList.Enabled = True
                vfgRecord.Editable = flexEDKbd
                mblnEdit = False
                vfgRecord.SelectionMode = flexSelectionByRow
                mlngResult = 0: Call vfgRecord_RowColChange
            End If
            
        Case 2:
            lngRetuId = mfrmReport.zlEditSave()
            If lngRetuId <> 0 Then
                mlngResult = lngRetuId:  Call zlRefList(mlngRecord, mlngResult)
                mintEditState = 0: Me.dtp日期.Enabled = True: Me.picLeft.Enabled = True: Me.rptList.SetFocus
            End If
        End Select
    Case conMenu_Edit_Untread:
        Select Case mintEditState   '当前编辑状态：0-非编辑状态,1-质控记录编辑,2-报告编辑
        Case 1
            If mblnEdit Then
                If MsgBox("是否放弃所做的修改！", vbInformation + vbOKCancel, Me.Caption) = vbCancel Then
                    Exit Sub
                Else
                    With vfgRecord
                        For lngRetuId = .FixedRows To .Rows - 1
                            If .TextMatrix(lngRetuId, mColL.结果) <> .TextMatrix(lngRetuId, mColL.原始结果) Then
                                .TextMatrix(lngRetuId, mColL.结果) = .TextMatrix(lngRetuId, mColL.原始结果)
                            End If
                        Next
                    End With
                End If
                mblnEdit = False
            End If
            vfgRecord.SelectionMode = flexSelectionByRow
            Me.PicList.Enabled = True
        Case 2: Call mfrmReport.zlEditCancel
             Me.picLeft.Enabled = True
        End Select
        
        mintEditState = 0: Me.dtp日期.Enabled = True
        Me.cbo科室.Enabled = True: Me.cbo仪器.Enabled = True
        
    Case conMenu_Edit_NewItem
        
        If frmQCTodayRecord.ZlEditStart(True, mlngRecord, mstrDate, mblnAllDev) <> 0 Then
            Call zlRefList
        End If
    Case conMenu_Edit_Modify

        mintEditState = 1: Me.dtp日期.Enabled = False: Me.PicList.Enabled = False
        Me.cbo科室.Enabled = False: Me.cbo仪器.Enabled = False
        vfgRecord.Editable = flexEDKbdMouse
        vfgRecord.SelectionMode = flexSelectionFree
    Case conMenu_Edit_Delete
        If mlngRecord = 0 Then Exit Sub
        With Me.rptList
            strInfo = "真的要删除该标本的质控登记，还原为普通标本吗？" & vbCrLf
            strInfo = strInfo & vbCrLf & "    标 本 号：" & .FocusedRow.Record(mCol.标本号).Value
            strInfo = strInfo & vbCrLf & "    检验仪器：" & .FocusedRow.Record(mCol.检验仪器).Value
            
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "Zl_检验质控记录_Edit(3," & mlngRecord & ")"
            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)

            Err = 0: On Error GoTo 0
            mlngRecord = 0: lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                If .Rows(lngRetuId + 1).GroupRow = False Then mlngRecord = .Rows(lngRetuId + 1).Record(mCol.标本ID).Value
            ElseIf lngRetuId > 0 Then
                If .Rows(lngRetuId - 1).GroupRow = False Then mlngRecord = .Rows(lngRetuId - 1).Record(mCol.标本ID).Value
            End If
            Call Me.zlRefList(mlngRecord)
        End With
    Case conMenu_Edit_Adjust                '填写失控报告
        If mfrmReport.ZlEditStart(mlngResult) Then
            mintEditState = 2: Me.dtp日期.Enabled = False: Me.picLeft.Enabled = False
        End If
    Case conMenu_Edit_Archive
        With Me.rptList
            strInfo = strInfo & vbCrLf & "    标 本 号：" & .FocusedRow.Record(mCol.标本号).Value
            strInfo = strInfo & vbCrLf & "    检验仪器：" & .FocusedRow.Record(mCol.检验仪器).Value
            strInfo = strInfo & vbCrLf & "    检验项目：" & Me.vfgRecord.TextMatrix(Me.vfgRecord.Row, mColL.中文名)
        End With
        
        If Me.vfgRecord.TextMatrix(Me.vfgRecord.Row, mColL.归档人) = "" Then
            strInfo = "真的要将当前失控报告归档吗？" & vbCrLf & strInfo
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "Zl_检验质控报告_Archive(" & mlngResult & ",0)"
            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            Me.vfgRecord.TextMatrix(Me.vfgRecord.Row, mColL.归档人) = UserInfo.姓名
        Else
            strInfo = "该失控报告已经归档，真的取消归档吗？" & vbCrLf & strInfo
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "Zl_检验质控报告_Archive(" & mlngResult & ",1)"
            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            Me.vfgRecord.TextMatrix(Me.vfgRecord.Row, mColL.归档人) = ""
        End If
        Call Me.zlRefList(mlngRecord, mlngResult)
    
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
    Case conMenu_View_Refresh        '刷新
        Call zlRefList(mlngRecord)
    
    Case conMenu_Tool_Analyse '计算
        With Me.vfgRecord
            lngRetuId = Val("" & .TextMatrix(.Row, mColL.项目id))
        End With
        If lngRetuId <= 0 Then
            MsgBox "请选择一个项目再使用此功能！", vbInformation, Me.Caption
            Exit Sub
        End If
        
        With Me.rptList
            If frmQCCompute.ShowMe(Me, _
                .FocusedRow.Record(mCol.仪器id).Value, lngRetuId, _
                CDate(mstrDate), .FocusedRow.Record(mCol.质控品id).Value) Then
                Call Me.zlRefList(mlngRecord)
            End If
        End With
    Case conMenu_Tool_Define '定值
        With Me.vfgRecord
            lngRetuId = Val("" & .TextMatrix(.Row, mColL.项目id))
        End With
        If lngRetuId <= 0 Then
            MsgBox "请选择一个项目再使用此功能！", vbInformation, Me.Caption
            Exit Sub
        End If
        With Me.rptList
            If .FocusedRow Is Nothing Then
                MsgBox "请选择一个质控品后再使用此功能！", vbInformation, Me.Caption
                Exit Sub
            End If
            If frmQCRedefine.ShowMe(Me, _
                .FocusedRow.Record(mCol.仪器id).Value, lngRetuId, _
                CDate(mstrDate), .FocusedRow.Record(mCol.质控品id).Value) Then
                Call Me.zlRefList(mlngRecord)
            End If
        End With
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else

            If Control.ID < conMenu_ReportPopup * 100# + 1 Or Control.ID > conMenu_ReportPopup * 100# + 99 Then Exit Sub

            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
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
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel: Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread: Control.Enabled = (mintEditState <> 0)
    
    Case conMenu_Edit_NewItem: Control.Enabled = (InStr(1, mstrPrivs, "登记") > 0 And mintEditState = 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Control.Enabled = (InStr(1, mstrPrivs, "登记") > 0 And mintEditState = 0 And mlngRecord <> 0)
        If Control.Enabled = False Then Exit Sub
'        Control.Enabled = (Trim(Me.rptList.FocusedRow.Record(mCol.归档人).Value) = "")
    Case conMenu_Edit_Adjust
        Control.Enabled = (InStr(1, mstrPrivs, "报告") > 0 And mintEditState = 0 And mlngResult <> 0)
        If Control.Enabled = False Then Exit Sub
'        Control.Enabled = (Val(Me.rptList.FocusedRow.Record(mCol.图标).Value) <> 0)
'        If Control.Enabled = False Then Exit Sub
'        Control.Enabled = (Trim(Me.rptList.FocusedRow.Record(mCol.归档人).Value) = "")
    Case conMenu_Edit_Archive
        Control.Enabled = (InStr(1, mstrPrivs, "归档") > 0 And mintEditState = 0 And mlngResult <> 0)
        If Control.Enabled = False Then Exit Sub
'        Control.Enabled = (Trim(Me.rptList.FocusedRow.Record(mCol.报告人).Value) <> "")
    
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    
    Case conMenu_Tool_Analyse
        Control.Enabled = (InStr(1, mstrPrivs, "计算") > 0 And mintEditState = 0 And mlngResult <> 0)
    Case conMenu_Tool_Define
        Control.Enabled = (InStr(1, mstrPrivs, "定值") > 0 And mintEditState = 0 And mlngResult <> 0)
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.picLeft.hWnd
    Case conPane_LJ
        Item.Handle = mfrmLJ.hWnd
    Case conPane_Report
        Item.Handle = mfrmReport.hWnd
    End Select
End Sub

Private Sub dtp日期_Change()
    Call zlRefList
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    
'    mstrPrivs = gstrPrivs
     '由于有技站要直接这个窗体所以重置一下脚本
    gstrPrivs = GetPrivFunc(100, 1210)
    mstrPrivs = gstrPrivs
    mblnAllDev = IIf(InStr(1, mstrPrivs, "所有科室") = 0, False, True)
    Me.cbo科室.Tag = "不刷新"
    Me.dtp日期.Value = Date: Me.dtp日期.MaxDate = Date
    mstrDate = Format(Date, "yyyy-MM-dd"): mlngRecord = 0: mlngResult = 0
    mintEditState = 0
    
    mlngEditWidth = Me.picLeft.Width
'    mlngEditHeight = frmQCTodayRecord.Height
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
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
'    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "登记(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "报告(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档(&T)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    cbrMenuBar.ID = xtpControlPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "失控计算(&Y)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "重新定值(&N)")
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
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "科室")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "科室")
    cbrCustom.Handle = Me.cbo科室.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "仪器")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "仪器")
    cbrCustom.Handle = Me.cbo仪器.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "检验日期")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "检验日期")
    cbrCustom.Handle = Me.dtp日期.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("Y"), conMenu_Tool_Analyse
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    Call zlDatabase.ShowReportMenu(Me.cbsThis, glngSys, glngModul, mstrPrivs)
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "登记"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "失控计算"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "重新定值")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "报告"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Set mfrmLJ = New frmQCChartLJ
    Set mfrmReport = New frmQCTodayReport
    
    Dim panThis As Pane, panSub As Pane
    Set panThis = dkpMan.CreatePane(conPane_List, 350, 400, DockLeftOf, Nothing)
    panThis.Title = "检验质控记录"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_LJ, 400, 600, DockRightOf, Nothing)
    panThis.Title = "仪器质控图形"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panSub = dkpMan.CreatePane(conPane_Report, 400, 200, DockBottomOf, panThis)
    panSub.Title = "仪器失控报告"
    panSub.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        .SetImageList Me.imgList
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 1024)   '必须在列设置之前设置，才能生效
        .AllowColumnRemove = False
        .AllowEdit = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False):  rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.标本ID, "标本ID", 0, False):  rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.标本号, "标本号", 62, False):  rptCol.Groupable = False: .SortOrder.Add rptCol
        
        Set rptCol = .Columns.Add(mCol.仪器id, "仪器id", 0, False):  rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.检验仪器, "检验仪器", 120, True):  rptCol.Groupable = True: rptCol.Visible = False: .GroupsOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.质控品id, "质控品id", 0, False):  rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.质控品, "质控品", 160, True):  rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.批号, "批号", 160, True):   rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.水平, "水平", 30, False):  rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.次数, "次数", 30, False):  rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
        .Populate
    End With
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    
    '装入基本数据
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHand
    
    If InStr(1, mstrPrivs, "所有科室") > 0 Then
        gstrSql = " Select Distinct b.Id, b.编码 , b.名称 As 科室 From 检验仪器 a ,部门表 b,检验质控品 c " & _
                  "Where a.使用小组ID = b.ID and a.id = c.仪器id order by b.编码 "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName)
        
    Else

        gstrSql = "Select Distinct B.ID, B.编码, B.名称 As 科室" & vbNewLine & _
                " From 检验仪器 A, 部门表 B, 检验质控品 C" & vbNewLine & _
                " Where A.使用小组id = B.ID And A.ID = C.仪器id And" & vbNewLine & _
                "      A.使用小组id In (Select Distinct D.使用小组id" & vbNewLine & _
                "                   From 检验小组成员 A, 检验小组 B, 检验小组仪器 C, 检验仪器 D" & vbNewLine & _
                "                   Where A.小组id = B.ID And B.ID = C.小组id　and 人员id = [1] And C.仪器id = D.ID)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, UserInfo.ID)
    End If
    
    Me.cbo科室.Clear
    If InStr(1, mstrPrivs, "所有科室") > 0 Then
        Me.cbo科室.AddItem "所有科室"
        Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = 0
    End If
    Do Until rsTemp.EOF
        Me.cbo科室.AddItem rsTemp("编码") & "-" & rsTemp("科室")
        Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = rsTemp("Id")
        rsTemp.MoveNext
    Loop
    If Me.cbo科室.ListCount = 0 Then MsgBox "尚未完成仪器使用小组的设置！", vbInformation, gstrSysName: Unload Me: Exit Sub
    Me.cbo科室.ListIndex = 0
    If Me.cbo科室.ListCount = 1 Then Me.cbo科室.Enabled = False
    Me.cbo科室.Tag = ""
    '数据装入
    Call zlRefList
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim panKind As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panKind = Me.dkpMan.FindPane(conPane_List)
'    panKind.MinTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, mlngEditHeight / Screen.TwipsPerPixelY
'    panKind.MaxTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, mlngEditHeight / Screen.TwipsPerPixelY
'    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panKind.MinTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, mlngEditHeight / Screen.TwipsPerPixelY
    panKind.MaxTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY
    Me.dkpMan.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmLJ
    Unload mfrmReport
    Set mfrmLJ = Nothing
    Set mfrmReport = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub fraNS_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        Me.fraNS.Top = Me.fraNS.Top + y
        Me.PicList.Height = Me.PicList.Height + y
        Me.vfgRecord.Top = Me.vfgRecord.Top + y
        Me.vfgRecord.Height = Me.vfgRecord.Height - y
    End If
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With Me.fraNS
        .Left = Me.picLeft.ScaleLeft: .Width = Me.picLeft.ScaleWidth - .Left
    End With
    With Me.vfgRecord
        .Left = Me.picLeft.ScaleLeft: .Width = Me.picLeft.ScaleWidth - .Left
        .Top = Me.fraNS.Top + Me.fraNS.Height
        .Height = Me.picLeft.ScaleHeight - .Top
    End With
    With Me.PicList
        .Left = Me.picLeft.ScaleLeft: .Width = Me.picLeft.ScaleWidth - .Left
        .Top = Me.picLeft.ScaleTop
        .Height = Me.picLeft.ScaleHeight - Me.vfgRecord.Height - Me.fraNS.Height
    End With
End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = Me.PicList.ScaleLeft: .Width = Me.PicList.ScaleWidth - .Left
        .Top = Me.PicList.ScaleTop
        .Height = Me.picLeft.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    
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

Private Sub rptList_SelectionChanged()
    Dim lng质控品id As Long, str日期 As String

    str日期 = Format(dtp日期.Value, "yyyy-MM-dd")
    If Me.rptList.FocusedRow Is Nothing Then
        mlngRecord = 0: mlngResult = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngRecord = 0: mlngResult = 0
    Else
        mlngRecord = Me.rptList.FocusedRow.Record.Item(mCol.标本ID).Value
        lng质控品id = Me.rptList.FocusedRow.Record.Item(mCol.质控品id).Value
'        mlngResult = Me.rptList.FocusedRow.Record.Item(mCol.结果ID).Value
    End If

    Call LoadRecord(str日期, lng质控品id, mlngRecord)
    
End Sub


Private Sub LoadRecord(ByVal str日期 As String, lng质控品id As Long, lng标本ID)
    '显示质控标本明细
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    On Error GoTo errH

    With Me.vfgRecord
        .Redraw = flexRDNone
        .Clear
        .Cols = 18
        .Rows = .FixedRows
        .TextMatrix(0, mColL.图标) = "": .ColWidth(mColL.图标) = 180: .FixedAlignment(mColL.图标) = flexAlignGeneralCenter
        .TextMatrix(0, mColL.ID) = "": .ColWidth(mColL.ID) = 0: .ColHidden(mColL.ID) = True
        .TextMatrix(0, mColL.中文名) = "中文名": .ColWidth(mColL.中文名) = 1500: .FixedAlignment(mColL.中文名) = flexAlignLeftCenter
        .TextMatrix(0, mColL.英文名) = "英文名": .ColWidth(mColL.英文名) = 800: .FixedAlignment(mColL.英文名) = flexAlignLeftCenter
        .TextMatrix(0, mColL.结果) = "结果值": .ColWidth(mColL.结果) = 900: .FixedAlignment(mColL.结果) = flexAlignRightCenter
        .TextMatrix(0, mColL.靶值) = "靶值": .ColWidth(mColL.靶值) = 800: .FixedAlignment(mColL.靶值) = flexAlignRightCenter
        .TextMatrix(0, mColL.SD) = "SD": .ColWidth(mColL.SD) = 800: .FixedAlignment(mColL.SD) = flexAlignRightCenter
        .TextMatrix(0, mColL.单位) = "单位": .ColWidth(mColL.单位) = 900: .FixedAlignment(mColL.单位) = flexAlignLeftCenter
        .TextMatrix(0, mColL.序号) = "序号": .ColWidth(mColL.序号) = 0: .ColHidden(mColL.序号) = True
        .TextMatrix(0, mColL.取值序列) = "取值序列": .ColWidth(mColL.取值序列) = 0: .ColHidden(mColL.取值序列) = True
        .TextMatrix(0, mColL.弃用结果) = "弃用结果": .ColWidth(mColL.弃用结果) = 0: .ColHidden(mColL.弃用结果) = True
        .TextMatrix(0, mColL.项目id) = "项目id": .ColWidth(mColL.项目id) = 0: .ColHidden(mColL.项目id) = True
        .TextMatrix(0, mColL.质控品id) = "质控品id": .ColWidth(mColL.质控品id) = 0: .ColHidden(mColL.质控品id) = True
        .TextMatrix(0, mColL.开始日期) = "开始日期": .ColWidth(mColL.开始日期) = 0: .ColHidden(mColL.开始日期) = True
        .TextMatrix(0, mColL.结束日期) = "结束日期": .ColWidth(mColL.结束日期) = 0: .ColHidden(mColL.结束日期) = True
        .TextMatrix(0, mColL.原始结果) = "原始结果": .ColWidth(mColL.原始结果) = 0: .ColHidden(mColL.原始结果) = True
        .TextMatrix(0, mColL.归档人) = "归档人": .ColWidth(mColL.归档人) = 0: .ColHidden(mColL.归档人) = True
        .TextMatrix(0, mColL.标记) = "标记": .ColWidth(mColL.标记) = 0: .ColHidden(mColL.标记) = True
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        
        strsql = "Select r.id,r.检验项目id, Nvl(f.标记, 0) As 标记, i.中文名, i.英文名, r.检验结果, x.均值, x.Sd, i.单位," & vbNewLine & _
                "            Decode(p.结果类型, 3, p.取值序列, '') As 取值序列, decode(p.排列序号,Null,i.编码,p.排列序号) As 序号,Nvl(r.弃用结果, 0) As 弃用结果," & vbNewLine & _
                "            x.质控品id,x.开始日期,x.结束日期,F.归档人,F.报告人 " & vbNewLine & _
                "From 检验普通结果 r, 检验质控报告 f," & vbNewLine & _
                "        (Select x.质控品id,x.项目id, x.均值, x.Sd,x.开始日期,nvl(x.结束日期,M.结束日期) as 结束日期 " & vbNewLine & _
                "            From 检验质控均值 x,检验质控品 M " & vbNewLine & _
                "            Where x.质控品id=M.id And x.质控品id = [2] And To_Date([1], 'yyyy-mm-dd') Between x.开始日期 And Nvl(x.结束日期, Sysdate)) x," & vbNewLine & _
                "        诊治所见项目 i, 检验项目 p" & vbNewLine & _
                "Where Nvl(r.弃用结果, 0) = 0 And r.Id = f.结果id(+) And r.检验标本id = [3] And r.检验项目id = x.项目id(+) And" & vbNewLine & _
                "           r.检验项目id = i.Id And r.检验项目id = p.诊治项目id" & vbNewLine & _
                "Order By decode(p.排列序号,Null,i.编码,p.排列序号)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, str日期, lng质控品id, lng标本ID)
        Do Until rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mColL.ID) = Val("" & rsTmp!ID)
            .TextMatrix(.Rows - 1, mColL.中文名) = Trim("" & rsTmp!中文名)
            .TextMatrix(.Rows - 1, mColL.英文名) = Trim("" & rsTmp!英文名)
            .TextMatrix(.Rows - 1, mColL.结果) = IIf(Left(Trim("" & rsTmp!检验结果), 1) = ".", "0" & Trim("" & rsTmp!检验结果), Trim("" & rsTmp!检验结果))
            .TextMatrix(.Rows - 1, mColL.原始结果) = .TextMatrix(.Rows - 1, mColL.结果)
            .TextMatrix(.Rows - 1, mColL.靶值) = IIf(Left(Trim("" & rsTmp!均值), 1) = ".", "0" & Trim("" & rsTmp!均值), Trim("" & rsTmp!均值))
            .TextMatrix(.Rows - 1, mColL.SD) = IIf(Left(Trim("" & rsTmp!SD), 1) = ".", "0" & Trim("" & rsTmp!SD), Trim("" & rsTmp!SD))
            .TextMatrix(.Rows - 1, mColL.单位) = Trim("" & rsTmp!单位)
            .TextMatrix(.Rows - 1, mColL.序号) = Trim("" & rsTmp!序号)
            .TextMatrix(.Rows - 1, mColL.取值序列) = Trim("" & rsTmp!取值序列)
            .TextMatrix(.Rows - 1, mColL.弃用结果) = Trim("" & rsTmp!弃用结果)
            .TextMatrix(.Rows - 1, mColL.项目id) = Val("" & rsTmp!检验项目id)
            .TextMatrix(.Rows - 1, mColL.质控品id) = Val("" & rsTmp!质控品id)
            .TextMatrix(.Rows - 1, mColL.开始日期) = Trim(Format("" & rsTmp!开始日期, "yyyy-MM-dd"))
            .TextMatrix(.Rows - 1, mColL.结束日期) = Trim(Format("" & rsTmp!结束日期, "yyyy-MM-dd"))
            .TextMatrix(.Rows - 1, mColL.归档人) = Trim("" & rsTmp!归档人)
            .TextMatrix(.Rows - 1, mColL.标记) = Trim("" & rsTmp!标记)
            If rsTmp!标记 <> 0 Then
                .Cell(flexcpBackColor, .Rows - 1, mColL.结果) = &HC0C0FF
                .Cell(flexcpFontBold, .Rows - 1, mColL.结果) = True
            End If
            rsTmp.MoveNext
        Loop
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then .Row = .FixedRows: .Col = 0
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vfgRecord_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vfgRecord
        If mblnEdit = False Then
            If Trim(.TextMatrix(Row, mColL.结果)) <> Trim(.TextMatrix(Row, mColL.原始结果)) Then
                mblnEdit = True
            End If
        End If
    End With
End Sub

Private Sub vfgRecord_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mintEditState <> 1 Then
        Cancel = True
        Exit Sub
    End If
    If Col <> mColL.结果 Then
        Cancel = True
        Exit Sub
    End If
    If Row < vfgRecord.FixedRows Then
        Cancel = True
        Exit Sub
    End If
    
End Sub

Private Sub vfgRecord_DblClick()
   If mlngRecord = 0 Then Exit Sub
    
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub vfgRecord_RowColChange()
    Dim lng项目ID As Long, lng质控品id As Long, str质控期间 As String, str开始日期 As String, str结束日期 As String
    If Me.cbo科室.Tag <> "" Then Exit Sub
    If mintEditState <> 0 Then Exit Sub
    With vfgRecord
        If mlngResult <> Val(.TextMatrix(.Row, mColL.ID)) And Val(.TextMatrix(.Row, mColL.ID)) <> 0 Then
            mlngResult = Val(.TextMatrix(.Row, mColL.ID))
            lng项目ID = Val(.TextMatrix(.Row, mColL.项目id))
            lng质控品id = Val(.TextMatrix(.Row, mColL.质控品id))
            
            str开始日期 = Format(dtp日期.Value, "yyyy-MM") & "-01"
            str结束日期 = Format(DateAdd("m", 1, CDate(Format(dtp日期.Value, "yyyy-MM") & "-01")) - 1, "yyyy-MM-dd")
            
            str质控期间 = lng质控品id & "=" & .TextMatrix(.Row, mColL.开始日期) & "," & .TextMatrix(.Row, mColL.结束日期)
                    
            Call mfrmReport.zlRefresh(mlngResult)
            Call mfrmLJ.zlRefresh(CStr(lng质控品id), lng项目ID, str开始日期, str结束日期, str质控期间)
            On Error Resume Next
            .SetFocus
            .Select .Row, mColL.结果
            
        End If
    End With
End Sub

Private Function zlEditSave() As Long
    '保存修改结果
    Dim strsql As String, rsTmp As ADODB.Recordset
    Dim lng仪器ID As Long, int性别 As Integer, str生日  As String
    Dim strItem As String, intRow As Integer, lng项目ID As Long
    
    If mblnEdit = False Then Exit Function
    strItem = ""
    With Me.vfgRecord
        For intRow = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(intRow, mColL.结果)) <> Trim(.TextMatrix(intRow, mColL.原始结果)) Then
                lng项目ID = Val(.TextMatrix(intRow, mColL.项目id))
                If lng项目ID <> 0 Then
                     .TextMatrix(intRow, mColL.原始结果) = Trim(.TextMatrix(intRow, mColL.结果))
                    strItem = strItem & "|" & lng项目ID & "^" & Trim(.TextMatrix(intRow, mColL.结果))
                End If
            End If
        Next
    End With
    If strItem <> "" Then
        strItem = Mid(strItem, 2)
        
        strsql = "Select 仪器id,标本类型,Decode(性别,'男',1,'女',2,0) as 性别,to_char(出生日期,'yyyy-MM-dd') as 生日 From 检验标本记录 where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mlngRecord)
        Do Until rsTmp.EOF
            '检验标本id_In ,仪器id_In ,标本类型_In,性别_in,出生日期_in,检验指标_in(项目ID^值|。。。) ,[微生物_in],[酶标板id_in]
            strsql = "Zl_检验普通结果_Batchupdate(" & mlngRecord & "," & rsTmp!仪器id & ",'" & rsTmp!标本类型 & "'," & rsTmp!性别 & _
                     IIf(Trim("" & rsTmp!生日) = "", ",Null", ",To_Date('" & rsTmp!生日 & "','yyyy-MM-dd')") & ",'" & strItem & "')"
            zlDatabase.ExecuteProcedure strsql, Me.Caption
            rsTmp.MoveNext
        Loop
        zlEditSave = mlngRecord
    End If
    
    
End Function


