VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRModelMan 
   Caption         =   "病历范文管理"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   Icon            =   "frmEPRModelMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicFile 
      BorderStyle     =   0  'None
      Height          =   4785
      Left            =   75
      ScaleHeight     =   4785
      ScaleWidth      =   2565
      TabIndex        =   8
      Top             =   705
      Width           =   2565
      Begin XtremeReportControl.ReportControl rptFile 
         Height          =   4800
         Left            =   0
         TabIndex        =   11
         Top             =   405
         Width           =   2445
         _Version        =   589884
         _ExtentX        =   4313
         _ExtentY        =   8467
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   645
         TabIndex        =   9
         Top             =   15
         Width           =   1725
      End
      Begin VB.Label lblFind 
         Caption         =   "查找(&V)"
         Height          =   405
         Left            =   0
         TabIndex        =   10
         Top             =   30
         Width           =   945
      End
   End
   Begin VB.PictureBox picNote 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2745
      ScaleHeight     =   345
      ScaleWidth      =   7515
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   750
      Width           =   7515
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明: "
         Height          =   180
         Left            =   90
         TabIndex        =   6
         Top             =   75
         Width           =   540
      End
   End
   Begin VB.PictureBox picTerm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   7815
      ScaleHeight     =   3150
      ScaleWidth      =   2445
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1290
      Width           =   2445
      Begin VSFlex8Ctl.VSFlexGrid vfgTerm 
         Height          =   2895
         Left            =   15
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   30
         Width           =   2340
         _cx             =   4128
         _cy             =   5106
         Appearance      =   2
         BorderStyle     =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483643
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   2
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
   Begin VB.PictureBox picList 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   2715
      ScaleHeight     =   3555
      ScaleWidth      =   4770
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1125
      Width           =   4770
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4215
         Left            =   915
         TabIndex        =   7
         Top             =   195
         Width           =   3375
         _Version        =   589884
         _ExtentX        =   5953
         _ExtentY        =   7435
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   240
         Top             =   2955
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":0B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":10BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":1458
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":1D32
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":20CC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vgdList 
         Height          =   900
         Left            =   3930
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2715
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6750
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRModelMan.frx":2466
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15346
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
   Begin MSComctlLib.ImageList imgFile 
      Left            =   360
      Top             =   5655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":2CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":3292
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":382C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":3DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":4360
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":48FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":4E94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1395
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   315
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRModelMan.frx":542E
      Left            =   930
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRModelMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const con_UnDefine = -999
Private Enum mPan
    File = 201
    Note = 202
    List = 203
    Term = 204
    View = 205
End Enum
Private Enum mFCol
    图标 = 0: ID: 种类: 编号: 名称: 保留: 简码
End Enum
Private Enum mLCol
    图标 = 0: 性质: ID: 分类: 编号: 名称: 简码: 说明: 部门: 人员
End Enum

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串
Private mstrKinds As String     '当前允许定义的病历类型串
Private mintPower As Integer    '示范管理权范围
'    mintPower=con_UnDefine，未定义;
'    mintPower=-1，不具备管理权;
'    mintPower=0，全院，这时显示所有的示范，也可以更改;
'    mintPower=1，科室，这时显示全院通用示范(科室id is null)和所在科室公有或部门内人员私有的示范，但不能更改全院通用示范;
'    mintPower=2，个人，这时显示全院通用示范(科室id is null)和所在科室通用示范(人员id is null)和个人示范，仅个人示范可更改

Private mlngFileID As Long      '当前文件ID
Private mblnShowAll As Boolean
Private WithEvents mfrmContent As frmEPRFileContent     '病历提纲窗格
Attribute mfrmContent.VB_VarHelpID = -1
Private mObjTabEpr As cTableEPR
'-----------------------------------------------------
'报表变量：用于发布到模块的报表使用
'-----------------------------------------------------
Private mlng编号 As Long
Private mstr名称 As String

'-----------------------------------------------------
'搜索框变量：用于控制搜索定位功能
'-----------------------------------------------------
Private mblnFindTag As Boolean      '搜索框焦点判断
Private mintLastRows As Integer     '搜索最后定位行位置

'-----------------------------------------------------
'以下为窗体公共方法
'-----------------------------------------------------
Public Sub RefreshList()
    '功能：刷新当前文档的内容，用于文档对象保存时执行的刷新处理
Dim lngItemID As Long
Dim lngCount As Long
    If Me.rptList.FocusedRow Is Nothing Then
        lngItemID = 0
    Else
        lngItemID = Me.rptList.FocusedRow.Record(mLCol.ID).Value
    End If
    lngCount = zlRefresh(mlngFileID, lngItemID)
    Me.stbThis.Panels(2).Text = "该种文件有" & lngCount & "个示范"
End Sub

Public Function zlRefFile(Optional lngFileID As Long) As Long
    '功能：刷新装入指定种类的病历文件清单，并定位到指定的文件上
Dim strGroups As String
Dim rsTemp As New ADODB.Recordset
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow
    
    mlng编号 = 0: mstr名称 = ""
    
    gstrSQL = "Select Id, 种类, 编号, 名称, 说明,保留" & vbNewLine & _
            "From 病历文件列表" & vbNewLine & _
            "Where 种类 In (" & mstrKinds & ") And Nvl(保留, 0) > = 0 And (种类 = 7 And 通用 > 0 Or 种类 <> 7" & IIf(mblnShowAll, "", " And 通用 > 0") & ")"
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Me.rptFile.Tag = ""
    Me.rptFile.Records.DeleteAll
    With rsTemp
        strGroups = ""
        Do While Not .EOF
            If InStr(1, strGroups, !种类) = 0 Then strGroups = strGroups & "," & !种类
            Set rptRcd = Me.rptFile.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr(!种类)): rptItem.Icon = rptItem.Value - 1
            rptRcd.AddItem CStr(!ID)
            Select Case !种类
            Case 1: rptRcd.AddItem CStr("1-门诊病历")
            Case 2: rptRcd.AddItem CStr("2-住院病历")
            Case 3: rptRcd.AddItem CStr("3-护理记录")
            Case 4: rptRcd.AddItem CStr("4-护理病历")
            Case 5: rptRcd.AddItem CStr("5-疾病证明报告")
            Case 6: rptRcd.AddItem CStr("6-知情文件")
            Case 7: rptRcd.AddItem CStr("7-诊疗报告")
            Case Else: rptRcd.AddItem ""
            End Select
            rptRcd.AddItem Val(CStr(!编号))
            rptRcd.AddItem CStr(!名称)
            rptRcd.AddItem NVL(!保留, 0)
            rptRcd.AddItem zl9ComLib.zlStr.PinYinCode(CStr(!名称))
            rptRcd.Tag = CStr("" & !说明)
            .MoveNext
        Loop
        If strGroups <> "" Then strGroups = Mid(strGroups, 2)
    End With
    With Me.rptFile
        If UBound(Split(strGroups, ",")) < 1 Then
            .GroupsOrder.DeleteAll
        ElseIf .GroupsOrder.Count = 0 Then
            .GroupsOrder.Add .Columns.Find(mFCol.种类)
            .GroupsOrder(0).SortAscending = True
        End If
        .Populate
    End With
    
    If lngFileID <> 0 Then
        For Each rptRow In Me.rptFile.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mFCol.ID).Value) = lngFileID Then
                    Set Me.rptFile.FocusedRow = rptRow: Exit For
                End If
            End If
        Next
    End If
    If Me.rptFile.Rows.Count > 0 Then
        If Me.rptFile.FocusedRow Is Nothing Then Set Me.rptFile.FocusedRow = Me.rptFile.Rows(0)
        If Me.rptFile.FocusedRow.GroupRow Then
            lngFileID = 0
        Else
            lngFileID = Me.rptFile.FocusedRow.Record.Item(mFCol.ID).Value
        End If
    Else
        lngFileID = 0
    End If
    
    zlRefFile = Me.rptFile.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefFile = Me.rptFile.Records.Count
    lngFileID = 0
End Function

Public Function zlRefresh(ByVal lngFileID As Long, Optional ByVal lngDemoId As Long) As Long
    '功能：刷新装入指定文件的示范目录
    '参数： lngFileId，文件ID
    '       lngDemoID，需要定位到的示范
    '返回：刷新装入的示范数目
Dim rsTemp As New ADODB.Recordset
Dim objItem As ReportRecordItem
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow
    
    Me.Tag = "zlRefresh"
    Err = 0: On Error GoTo errHand
    Select Case mintPower
    Case 0
        gstrSQL = "Select l.Id, l.编号, l.名称, l.简码, Nvl(l.分类,'未分类') As 分类,l.性质, l.说明, l.通用级, d.名称 As 部门, p.姓名 As 人员,Decode(l.分类,Null,1,2) As 排序 " _
                & "From 病历范文目录 l, 部门表 d, 人员表 p " _
                & "Where l.科室id = d.Id And l.人员id = p.Id And l.文件id =[1] " _
                & "Order By Decode(l.分类,Null,1,2),l.分类,l.编号"
    Case 1
        gstrSQL = "Select l.Id, l.编号, l.名称, l.简码, Nvl(l.分类,'未分类') As 分类,l.性质, l.说明, l.通用级, d.名称 As 部门, p.姓名 As 人员,Decode(l.分类,Null,1,2) As 排序 " _
                & "From 病历范文目录 l, 部门表 d, 人员表 p " _
                & "Where l.科室id = d.Id(+) And l.人员id = p.Id(+) And l.文件id =[1] And " _
                & "      (Nvl(l.通用级, 0) = 0 Or " _
                & "      l.通用级 in (1,2) And l.科室id In (Select r.部门id From 部门人员 r, 上机人员表 u Where r.人员id = u.人员id And u.用户名 = User)) " _
                & "Order By Decode(l.分类,Null,1,2),l.分类,l.编号"
    Case Else
        gstrSQL = "Select l.Id, l.编号, l.名称, l.简码, Nvl(l.分类,'未分类') As 分类,l.性质, l.说明, l.通用级, d.名称 As 部门, p.姓名 As 人员,Decode(l.分类,Null,1,2) As 排序 " _
                & "From 病历范文目录 l, 部门表 d, 人员表 p " _
                & "Where l.科室id = d.Id(+) And l.人员id = p.Id(+) And l.文件id =[1] And " _
                & "      (Nvl(l.通用级, 0) = 0 Or " _
                & "      l.通用级 =1 And l.科室id In (Select r.部门id From 部门人员 r, 上机人员表 u Where r.人员id = u.人员id And u.用户名 = User) Or " _
                & "      l.通用级 =2 And l.人员id In (Select u.人员id From 上机人员表 u Where u.用户名 = User)) " _
                & "Order By Decode(l.分类,Null,1,2),l.分类,l.编号"
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    
    Me.rptList.Records.DeleteAll
    Do While Not rsTemp.EOF
        Set rptRcd = Me.rptList.Records.Add()
        Set rptItem = rptRcd.AddItem(CInt(IIf(IsNull(rsTemp!通用级), 0, rsTemp!通用级))): rptItem.Icon = rptItem.Value
        Set rptItem = rptRcd.AddItem(CInt(Val("" & rsTemp!性质))): rptItem.Icon = IIf(rptItem.Value = 0, 4, IIf(rptItem.Value = 1, 5, 3))
        rptRcd.AddItem CStr(rsTemp!ID)
                
        Set objItem = rptRcd.AddItem(Val(rsTemp!排序) & CStr(rsTemp!分类))
        objItem.Caption = CStr(rsTemp!分类)
                        
        rptRcd.AddItem ZLCommFun.NVL(rsTemp!编号)
        rptRcd.AddItem CStr(rsTemp!名称)
        rptRcd.AddItem CStr("" & rsTemp!简码)
        rptRcd.AddItem CStr("" & rsTemp!说明)
        rptRcd.AddItem CStr("" & rsTemp!部门)
        rptRcd.AddItem CStr("" & rsTemp!人员)
        rsTemp.MoveNext
    Loop
    Me.rptList.Populate
    
    If Me.rptList.Rows.Count > 0 Then
        For Each rptRow In Me.rptList.Rows
            If Not (rptRow.Record Is Nothing) Then

                If lngDemoId = rptRow.Record(mLCol.ID).Value Then Set Me.rptList.FocusedRow = rptRow: Exit For
            
            End If
        Next
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Me.Tag = ""
    Call rptList_SelectionChanged
    zlRefresh = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefresh = Me.rptList.Records.Count
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
    If Me.rptFile.FocusedRow Is Nothing Then
        strSubhead = ""
    ElseIf Me.rptFile.FocusedRow.GroupRow Then
        strSubhead = ""
    Else
        strSubhead = Me.rptFile.FocusedRow.Record(mFCol.名称).Value
    End If
    
    Set objPrint.Body = Me.vgdList
    objPrint.Title.Text = strSubhead & "示范目录"
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
'以下为窗体控件方法
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim lngDemoId As Long
Dim cbrControl As CommandBarControl
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_ExportToXML:
        '导出到XML文件
        If Me.rptFile.FocusedRow Is Nothing Then Exit Sub
        If Me.rptFile.FocusedRow.GroupRow = True Then Exit Sub
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        
        Dim strF As String
        lngDemoId = Me.rptList.FocusedRow.Record.Item(mLCol.ID).Value
        '普通住院病历
        dlgThis.Filename = "示范_" & Me.rptFile.FocusedRow.Record.Item(mFCol.名称).Value & "_" & Me.rptList.FocusedRow.Record.Item(mLCol.名称).Value & ".xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        Err = 0: On Error Resume Next
        dlgThis.ShowSave
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        Err = 0: On Error GoTo 0
        On Error GoTo errHand
        strF = dlgThis.Filename
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("该文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        If rptList.FocusedRow.Record(mLCol.性质).Value = 2 Then '表格式编辑器
            mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_全文示范编辑, lngDemoId, False, 0
            If mObjTabEpr.zlExportXML(strF) Then
                MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        Else
            Dim DocXML As New cEPRDocument
            DocXML.InitEPRDoc cprEM_修改, cprET_全文示范编辑, lngDemoId
            DocXML.KeepRTF = True
            DocXML.OpenEPRDoc DocXML.frmEditor.Editor1
            If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                DoEvents
                MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        End If
    Case conMenu_File_ExportToXMLs
        frmModelExportOrImport.ShowMe Me, 1
    Case conMenu_File_ImportFromXMLs
        frmModelExportOrImport.ShowMe Me, 2
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_NewItem
        If mlngFileID = 0 Then Exit Sub
        lngDemoId = frmEPRModelEdit.ShowMe(Me, True, CByte(mintPower), mlngFileID, 0, rptFile.FocusedRow.Record.Item(mFCol.保留).Value)
        If lngDemoId <> 0 Then
            Call Me.zlRefresh(mlngFileID, lngDemoId)
            Me.stbThis.Panels(2).Text = "该种文件现有" & Me.rptList.Rows.Count & "个示范"
        End If
    Case conMenu_Edit_Modify
        If mlngFileID = 0 Then Exit Sub
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        If Me.rptList.FocusedRow.Record Is Nothing Then Exit Sub
        
        lngDemoId = Me.rptList.FocusedRow.Record.Item(mLCol.ID).Value
        lngDemoId = frmEPRModelEdit.ShowMe(Me, False, CByte(mintPower), mlngFileID, lngDemoId)
        If lngDemoId <> 0 Then Call Me.zlRefresh(mlngFileID, lngDemoId)
    Case conMenu_Edit_Delete
        Dim lngIndex As Long, strMsg As String
        With Me.rptList
            If .FocusedRow Is Nothing Then Exit Sub
            strMsg = "真的删除该示范吗？" & vbCrLf & "——" & .FocusedRow.Record(mLCol.名称).Value
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "zl_病历范文目录_delete('" & .FocusedRow.Record(mLCol.ID).Value & "')"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            Err = 0: On Error GoTo 0
            lngIndex = .FocusedRow.Record.Index
            Call .Records.RemoveAt(.FocusedRow.Record.Index)
            .Populate
            If .Records.Count <> 0 Then
                If lngIndex >= .Records.Count Then lngIndex = 0
                lngDemoId = .Records(lngIndex).Item(mLCol.ID).Value
            Else
                lngDemoId = 0
            End If
            Call Me.zlRefresh(mlngFileID, lngDemoId)
            Me.stbThis.Panels(2).Text = "该种文件剩余" & Me.rptList.Rows.Count & "个示范"
        End With
    Case conMenu_Edit_Compend
        If Me.rptFile.FocusedRow Is Nothing Then Exit Sub
        If Me.rptFile.FocusedRow.GroupRow = True Then Exit Sub
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        lngDemoId = Me.rptList.FocusedRow.Record(mLCol.ID).Value
        If rptList.FocusedRow.Record(mLCol.性质).Value = 2 Then '表格式编辑器
            On Error GoTo errHand
            mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_全文示范编辑, lngDemoId
        Else
            Dim Doc As New cEPRDocument
            Doc.InitEPRDoc cprEM_修改, cprET_全文示范编辑, lngDemoId
            Doc.ShowEPREditor Me
        End If
    Case conMenu_Edit_Request
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        lngDemoId = Me.rptList.FocusedRow.Record.Item(mLCol.ID).Value
        If frmEPRModelRequest.ShowMe(Me, lngDemoId, mintPower) = True Then Call rptList_SelectionChanged
    
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Option
        mblnShowAll = Not mblnShowAll
        Control.Checked = mblnShowAll
        Call zlRefFile
    Case conMenu_View_LocationItem
        txtFind.SetFocus
    Case conMenu_View_Refresh
        Call zlRefFile(mlngFileID)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
        '执行发布到当前模块的报表
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If mstr名称 <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                    "文件ID=" & mlngFileID, "编号=" & mlng编号, "名称=" & mstr名称)
            Else
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            End If
        End If
    End Select
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    If mblnFindTag = True Then
        txtFind.ForeColor = vbBlack
        If txtFind.Text = "请输入名称或拼音简码" Then txtFind.Text = ""
    Else
        If txtFind.Text = "" Then txtFind.ForeColor = vbGrayText: txtFind.Text = "请输入名称或拼音简码"
    End If
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = (mintPower >= 0)
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptFile.Records.Count <> 0)
    Case conMenu_Edit_NewItem
        Control.Visible = (mintPower >= 0)
        Control.Enabled = (mlngFileID <> 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Request
    
        Control.Visible = (mintPower >= 0)
        
        Control.Enabled = True
        If Me.rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        ElseIf Me.rptList.FocusedRow.Record Is Nothing Then
            Control.Enabled = False
        Else
            If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record.Item(mLCol.图标).Value >= mintPower)
        End If

    Case conMenu_Edit_Compend
        Control.Visible = (mintPower >= 0)
        Control.Enabled = True
        If Me.rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        ElseIf Me.rptList.FocusedRow.Record Is Nothing Then
            Control.Enabled = False
        Else
            If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record.Item(mLCol.图标).Value >= mintPower)
        End If
    Case conMenu_File_ExportToXML
        Control.Enabled = True
        If Me.rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        ElseIf Me.rptList.FocusedRow.Record Is Nothing Then
            Control.Enabled = False
        Else
            If Control.Enabled Then Control.Enabled = Not Me.rptFile.FocusedRow.GroupRow
            If Control.Enabled Then Control.Enabled = Not (Me.rptList.FocusedRow Is Nothing)
        End If
        
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Option: Control.Checked = mblnShowAll
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPan.File
        Item.Handle = Me.PicFile.hWnd
    Case mPan.Note
        Item.Handle = Me.picNote.hWnd
    Case mPan.List
        Item.Handle = Me.picList.hWnd
    Case mPan.Term
        Item.Handle = Me.picTerm.hWnd
    Case mPan.View
        Item.Handle = mfrmContent.hWnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
Dim rptCol As ReportColumn
Dim lngCount As Long
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    mstrKinds = ""
    If InStr(1, mstrPrivs, "门诊病历范文") > 0 Then mstrKinds = mstrKinds & ",1"
    If InStr(1, mstrPrivs, "住院病历范文") > 0 Then mstrKinds = mstrKinds & ",2"
    If InStr(1, mstrPrivs, "护理病历范文") > 0 Then mstrKinds = mstrKinds & ",4"
    If InStr(1, mstrPrivs, "疾病证明报告范文") > 0 Then mstrKinds = mstrKinds & ",5"
    If InStr(1, mstrPrivs, "知情文件范文") > 0 Then mstrKinds = mstrKinds & ",6"
    If InStr(1, mstrPrivs, "诊疗报告范文") > 0 Then mstrKinds = mstrKinds & ",7"
    If mstrKinds <> "" Then mstrKinds = Mid(mstrKinds, 2)
    mblnShowAll = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ShowAll", False)
    
    Set mObjTabEpr = New cTableEPR
    mObjTabEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    
    If InStr(1, gstrPrivsEpr, "全院病历范文") <> 0 Then
        mintPower = 0
    ElseIf InStr(1, gstrPrivsEpr, "科室病历范文") <> 0 Then
        mintPower = 1
    ElseIf InStr(1, gstrPrivsEpr, "个人病历范文") <> 0 Then
        mintPower = 2
    Else
        mintPower = -1
    End If
    
    Call ZLCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = ZLCommFun.GetPubIcons
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
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出为XML文件(&L)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXMLs, "批量导出范文文件(&E)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ImportFromXMLs, "批量导入范文文件(&I)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "内容(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "条件(&Q)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "查找(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "显示未用病历(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("F"), conMenu_View_LocationItem
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("D"), conMenu_Edit_Compend
        .Add FCONTROL, Asc("R"), conMenu_Edit_Request
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "内容"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "条件")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '---------------------------------------------------------------
    '读取发布到该模块的报表:因为是一次性读取,全局变量可用
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panThis As Pane, panNote As Pane, panView As Pane, panList As Pane, panTerm As Pane
    If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
    
    Set panThis = dkpMan.CreatePane(mPan.File, 150, 480, DockLeftOf, Nothing)
    panThis.Title = "文件列表": panThis.Options = PaneNoCaption
    Set panNote = dkpMan.CreatePane(mPan.Note, 500, 25, DockRightOf, Nothing)
    panNote.Title = "文件说明": panNote.Options = PaneNoCaption
    Set panView = dkpMan.CreatePane(mPan.View, 500, 240, DockBottomOf, panNote)
    panView.Title = "范文内容": panView.Options = PaneNoCaption
    Set panList = dkpMan.CreatePane(mPan.List, 400, 240, DockTopOf, panView)
    panList.Title = "范文列表": panList.Options = PaneNoCaption
    Set panTerm = dkpMan.CreatePane(mPan.Term, 120, 240, DockRightOf, panList)
    panTerm.Title = "应用条件": panTerm.Options = PaneNoCaption
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptFile
        Set rptCol = .Columns.Add(mFCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mFCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mFCol.种类, "种类", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mFCol.编号, "编号", 49, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mFCol.名称, "名称", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mFCol.保留, "保留", 0, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mFCol.简码, "简码", 0, False): rptCol.Editable = False: rptCol.Groupable = False
        
        .SetImageList Me.imgFile
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
    With Me.rptList
        Set rptCol = .Columns.Add(mLCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mLCol.性质, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mLCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mLCol.分类, "分类", 0, False): rptCol.Editable = False: rptCol.Groupable = False:  rptCol.Visible = False: rptCol.Sortable = False
        Set rptCol = .Columns.Add(mLCol.编号, "编号", 49, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mLCol.名称, "名称", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mLCol.简码, "简码", 60, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mLCol.说明, "说明", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mLCol.部门, "部门", 70, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mLCol.人员, "编制人", 50, False): rptCol.Editable = False: rptCol.Groupable = True
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        
        .SetImageList Me.imgList

        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mLCol.分类)
        .GroupsOrder(0).SortAscending = True
        
        .SortOrder.Add .Columns.Find(mLCol.编号)
        
    End With
    
    '-----------------------------------------------------
    '查询框初始化
    mblnFindTag = False
    txtFind.ForeColor = vbGrayText
    txtFind.Text = "请输入名称或拼音简码"
    mintLastRows = 0
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '数据装入
    If mstrKinds = "" Then
        DoEvents
        Me.stbThis.Panels(2).Text = "你不具备任何种类的示范管理权限"
    Else
        lngCount = Me.zlRefFile()
        Me.stbThis.Panels(2).Text = "共有" & lngCount & "个文件"
    End If
End Sub

Private Sub Form_Resize()
    Dim panThis As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Set panThis = Me.dkpMan.FindPane(mPan.File)
    panThis.MinTrackSize.SetSize 3300 / Screen.TwipsPerPixelX, 0
    panThis.MaxTrackSize.SetSize 3300 / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panThis.MinTrackSize.SetSize 0, 0
    panThis.MaxTrackSize.SetSize 3300 / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    
    Set panThis = Me.dkpMan.FindPane(mPan.Note)
    panThis.MinTrackSize.SetSize 0, 345 / Screen.TwipsPerPixelY
    panThis.MaxTrackSize.SetSize panThis.MaxTrackSize.Width, 345 / Screen.TwipsPerPixelY
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmContent
    Set mfrmContent = Nothing
    Set mObjTabEpr = Nothing
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ShowAll", mblnShowAll
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mfrmContent_DblClick()
Dim cbrControl As CommandBarControl
    If mlngFileID = 0 Then Exit Sub
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Compend)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub PicFile_Resize()
    lblFind.Move 70, 90, lblFind.Width, lblFind.Height
    If PicFile.Width > 800 Then txtFind.Move 800, 50, PicFile.Width - 800, 300
    If PicFile.Height > 400 Then rptFile.Move 0, 400, PicFile.Width, PicFile.Height - 400
End Sub

Private Sub piclist_Resize()
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = 0: .Width = Me.picList.ScaleWidth
        .Top = 0: .Height = Me.picList.ScaleHeight
    End With
End Sub

Private Sub picTerm_Resize()
    Err = 0: On Error Resume Next
    With Me.vfgTerm
        .Left = 0: .Width = Me.picTerm.ScaleWidth
        .Top = 0: .Height = Me.picTerm.ScaleHeight
        .AutoSize 0
    End With
End Sub

Private Sub rptFile_SelectionChanged()
Dim lngCount As Long
    With Me.rptFile
        If .FocusedRow Is Nothing Then
            mlngFileID = 0: Me.lblNote.Caption = "说明:"
        ElseIf .FocusedRow.GroupRow = True Then
            mlngFileID = 0: Me.lblNote.Caption = "说明:"
        Else
            mlngFileID = .FocusedRow.Record.Item(mFCol.ID).Value
            Me.lblNote.Caption = "说明: " & .FocusedRow.Record.Tag
        End If
        mlng编号 = 0: mstr名称 = ""
    End With
    If Me.rptFile.FocusedRow Is Nothing Then Exit Sub
    If Me.rptFile.Tag = "" Or Val(Me.rptFile.Tag) <> Me.rptFile.FocusedRow.Index Then
        lngCount = zlRefresh(mlngFileID)
        Me.rptFile.Tag = Me.rptFile.FocusedRow.Index
        If lngCount = 0 Then
            mfrmContent.edtThis.ForceEdit = True
            mfrmContent.edtThis.ReadOnly = False
            mfrmContent.edtThis.NewDoc
            mfrmContent.edtThis.ReadOnly = True
            mfrmContent.edtThis.ForceEdit = False
        End If
        Me.stbThis.Panels(2).Text = "该种文件有" & lngCount & "个示范"
    End If
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(mLCol.编号))
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Dim cbrPopupBar As CommandBar
Dim cbrPopupItem As CommandBarControl
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup

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

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim cbrControl As CommandBarControl
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub rptList_SelectionChanged()
    Dim lngItemID As Long
    Dim rsTemp As New ADODB.Recordset
    
    '如果特殊处理过程导致的选择区域变化，如连续刷新过程，则直接退出
    If Me.Tag <> "" Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then
        lngItemID = 0
        mlng编号 = 0
        mstr名称 = ""
        Call mfrmContent.zlRefresh(0, cprEmCPKModelEssay)
    ElseIf Me.rptList.FocusedRow.Record Is Nothing Then
    
        lngItemID = 0
        mlng编号 = 0
        mstr名称 = ""
        Call mfrmContent.zlRefresh(0, cprEmCPKModelEssay)
    Else
    
        lngItemID = Me.rptList.FocusedRow.Record.Item(mLCol.ID).Value
        mlng编号 = Val(Me.rptList.FocusedRow.Record.Item(mLCol.编号).Value)
        mstr名称 = Me.rptList.FocusedRow.Record.Item(mLCol.名称).Value
    End If
    
    '刷新内容
    Call mfrmContent.zlRefresh(lngItemID, cprEmCPKModelEssay)
    
    '刷新条件
    Err = 0: On Error GoTo errHand
    Me.vfgTerm.Clear: Me.vfgTerm.Rows = Me.vfgTerm.FixedRows
    Set Me.vfgTerm.Cell(flexcpPicture, Me.vfgTerm.FixedRows - 1, 0) = Me.imgList.ListImages(4).Picture
    gstrSQL = "Select 名称 As 条件项, 简码 As 条件值" & vbNewLine & _
            "From Table(Cast(f_Segment_条件项([1]) As " & gstrDbOwner & ".t_Dic_Rowset))" & vbNewLine & _
            "Where 简码 Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngItemID)
    With rsTemp
        If .RecordCount <= 0 Then
            Me.vfgTerm.TextMatrix(Me.vfgTerm.FixedRows - 1, 0) = "无使用限制条件。"
        Else
            Me.vfgTerm.TextMatrix(Me.vfgTerm.FixedRows - 1, 0) = "在以下条件满足时可以使用："
        End If
        Do While Not .EOF
            Me.vfgTerm.Rows = Me.vfgTerm.Rows + 1
            Me.vfgTerm.TextMatrix(Me.vfgTerm.Rows - 1, 0) = Space(2) & Me.vfgTerm.Rows - 1 & ")" & !条件项 & "为'" & Replace(!条件值, vbTab, "'或'") & "'"
            .MoveNext
        Loop
    End With
    Me.vfgTerm.AutoSize 0
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtFind_Change()
    mintLastRows = 0
End Sub

Private Sub txtFind_GotFocus()
    mblnFindTag = True
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intCount As Integer

    If KeyCode = vbKeyReturn And txtFind.Text <> "" Then
        For intCount = mintLastRows + 1 To Me.rptFile.Rows.Count - 1
            If Me.rptFile.Rows(intCount).GroupRow = False Then
                If InStr(Me.rptFile.Rows(intCount).Record(mFCol.名称).Value, txtFind.Text) Or InStr(Me.rptFile.Rows(intCount).Record(mFCol.简码).Value, UCase(txtFind.Text)) Then
                    Set Me.rptFile.FocusedRow = Me.rptFile.Rows(intCount)
                    mintLastRows = intCount
                    Exit For
                End If
            End If
        Next
        If intCount = Me.rptFile.Rows.Count And mintLastRows = 0 Then
            Call MsgBox("未找到与“" & txtFind.Text & "”匹配的范文，请重新输入名称或简码。", vbInformation, gstrSysName)
            txtFind.Text = ""
        End If
    End If
    txtFind.SetFocus
End Sub

Private Sub txtFind_LostFocus()
    mblnFindTag = False
End Sub
