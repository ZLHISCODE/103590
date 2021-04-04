VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmAppRuleList 
   Caption         =   "仪器质控设置"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   Icon            =   "frmAppRuleList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   11670
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picRes 
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   4800
      ScaleHeight     =   3795
      ScaleWidth      =   6225
      TabIndex        =   4
      Top             =   1170
      Width           =   6225
      Begin XtremeReportControl.ReportControl rptRule 
         Height          =   3405
         Left            =   165
         TabIndex        =   5
         Top             =   420
         Width           =   5415
         _Version        =   589884
         _ExtentX        =   9551
         _ExtentY        =   6006
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnSort =   0   'False
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
   End
   Begin VB.PictureBox picDev 
      BorderStyle     =   0  'None
      Height          =   5370
      Left            =   135
      ScaleHeight     =   5370
      ScaleWidth      =   4575
      TabIndex        =   2
      Top             =   405
      Width           =   4575
      Begin XtremeReportControl.ReportControl rptDev 
         Height          =   4560
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4395
         _Version        =   589884
         _ExtentX        =   7752
         _ExtentY        =   8043
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7155
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppRuleList.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15505
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   510
      Top             =   5730
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleList.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleList.frx":13B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleList.frx":1950
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleList.frx":1EEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1260
      Left            =   2355
      TabIndex        =   1
      Top             =   4650
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
      Bindings        =   "frmAppRuleList.frx":2484
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmAppRuleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mColD
    图标 = 0: ID: 名称: 项目id: 编码: 项目: 英文名: 质控周期: 水平数
End Enum
Private Enum mColR
    性质 = 0: ID: 判断: 规则: 批范围: 多水平: 符合处理: 不符处理: 是否使用
End Enum

Const conPane_Dev = 201
Const conPane_Res = 202
Const conPane_Edit = 203

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串
Private mfrmEdit As frmAppRuleEdit
Private mLngEditWidth As Long, mLngEditHeight As Long

Private mintEditState As Integer    '当前编辑状态：0-非编辑状态,1-编辑状态
Private mlngDevId As Long           '仪器id
Private mlngRuleId As Long          '规则id
Private mblnStart As Boolean        '是否已经增加了多规则起点项目,在规则列表刷新时执行
Private mlngItemID As Long         '项目ID

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
Private Function zlRefDev() As Long
    '功能：刷新装入指定仪器
    Dim rsTemp As New ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    Dim rptItem As ReportRecordItem
    Dim objRow As ReportRow
    Dim lngYQ As Long
    gstrSql = "Select A.ID, A.编码, A.名称, D.名称 As 使用部门, Count(S.项目id) As 是否失控," & vbNewLine & _
            "       Decode(Nvl(A.质控周期, 0), 0, '', A.质控周期 || Nvl(A.周期单位, '天')) As 质控周期, A.质控水平数 As 水平数" & vbNewLine & _
            "From 检验仪器 A, 部门表 D, 检验仪器状态 S" & vbNewLine & _
            "Where A.使用小组id = D.ID(+) And A.ID = S.仪器id(+) And Nvl(A.微生物, 0) <> 1" & vbNewLine & _
            "Group By A.ID, A.编码, A.名称, D.名称, A.质控周期, A.周期单位, A.质控水平数"

    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.rptDev.Records.DeleteAll
    lngYQ = 0
    With rsTemp
        Do Until .EOF
            
            
            gstrSql = "Select Distinct C.项目ID,I.编码, I.名称 As 中文名, L.缩写 As 英文名" & vbNewLine & _
                "From 检验仪器项目 C, 检验项目 L, 检验报告项目 R, 诊疗项目目录 I, 检验质控品项目 A" & vbNewLine & _
                "Where A.项目id = C.项目id And C.项目id = L.诊治项目id And L.诊治项目id = R.报告项目id And R.诊疗项目id = I.ID And" & vbNewLine & _
                "      I.组合项目 <> 1 And L.项目类别 <> 2 And C.仪器id = [1]" & vbNewLine & _
                "Order By I.编码"
            Set rsItem = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val("" & !ID))
            Do Until rsItem.EOF
                
                Set rptRcd = Me.rptDev.Records.Add()
                Set rptItem = rptRcd.AddItem("0"): rptItem.Focusable = False
                If Val("" & !是否失控) = 0 Then
                    rptItem.Icon = 0
                Else
                    rptItem.Icon = 1
                End If
                Set rptItem = rptRcd.AddItem(CStr("" & !ID)): rptItem.Focusable = False
                Set rptItem = rptRcd.AddItem(CStr("" & !编码 & "-" & !名称 & "(" & !使用部门 & ")")): rptItem.Focusable = False
                Set rptItem = rptRcd.AddItem(CStr("" & rsItem!项目id)): rptItem.Focusable = False
                Set rptItem = rptRcd.AddItem(CStr("" & rsItem!编码)): rptItem.Focusable = False
                Set rptItem = rptRcd.AddItem(CStr("" & rsItem!中文名)): rptItem.Focusable = False
                Set rptItem = rptRcd.AddItem(CStr("" & rsItem!英文名)): rptItem.Focusable = False
                Set rptItem = rptRcd.AddItem(CStr("" & !质控周期)): rptItem.Focusable = False
                Set rptItem = rptRcd.AddItem(CStr("" & !水平数)): rptItem.Focusable = False
                
                rsItem.MoveNext
            Loop
            lngYQ = lngYQ + 1
            .MoveNext
        Loop
    End With
    


    
    Me.rptDev.Populate
    Call rptDev_SelectionChanged
    
    '折叠所有组
    For Each objRow In Me.rptDev.Rows
        If objRow.GroupRow Then objRow.Expanded = False
    Next
    
    If Me.rptDev.FocusedRow Is Nothing And Me.rptDev.Rows.Count > 0 Then
        If Me.rptDev.Rows(0).GroupRow Then
            Me.rptDev.Rows(0).Expanded = True
            Set Me.rptDev.FocusedRow = Me.rptDev.Rows(0).Childs(0)
        Else
            Set Me.rptDev.FocusedRow = Me.rptDev.Rows(0)
        End If
    End If
    
    zlRefDev = Me.rptDev.Records.Count
    Me.stbThis.Panels(2).Text = "共有" & lngYQ & "台仪器"
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefDev = Me.rptDev.Records.Count
End Function

Private Function zlRefRule(Optional lngRuleId As Long) As Long
    '功能：刷新装入当前仪器的规则
    Dim rsTemp As New ADODB.Recordset
    Dim objParent As Object, lngChilds As Long
    
    mblnStart = False
    
    gstrSql = "Select R.ID, R.上级id, R.性质, Decode(R.性质, 'Y', '符合: ', 'N', '不符: ', '') || R.判断 As 判断, B.名称 As 规则," & vbNewLine & _
            "       Decode(R.批范围, 1, '当前批', '近' || R.批范围 || '批') As 批范围, Decode(R.多水平, 1, '多', '') As 多水平," & vbNewLine & _
            "       Decode(Y结束, 0, '下一步', '结束') As 符合处理, Decode(N结束, 0, '下一步', '结束') As 不符处理, Decode(是否使用,1,'√','') as 是否使用" & vbNewLine & _
            "From (Select Level As 层次, ID, Nvl(上级id, 0) As 上级id, 性质, 判断, 规则id, 批范围, 多水平, Y结束, N结束, 是否使用" & vbNewLine & _
            "       From 检验仪器规则" & vbNewLine & _
            "       Where 仪器id = [1] And nvl(项目ID,0)=[2] " & vbNewLine & _
            "       Start With 仪器id = [1] And 上级id Is Null" & vbNewLine & _
            "       Connect By Prior ID = 上级id) R, 检验质控规则 B" & vbNewLine & _
            "Where R.规则id = B.ID" & vbNewLine & _
            "Order By R.层次, Decode(R.性质, '0', 0, '1', 1, 'Y', 2, 'N', 3, 1)"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevId, mlngItemID)
    Err = 0: On Error GoTo 0
    Me.rptRule.Records.DeleteAll
    Me.rptRule.Populate
    With rsTemp
        Do While Not .EOF
            If Val("" & !上级ID) = 0 Then
                Set rptRcd = Me.rptRule.Records.Add()
            Else
                Me.rptRule.Populate
                For Each rptRow In Me.rptRule.Rows
                    If Val(rptRow.Record(mColR.ID).Value) = Val("" & !上级ID) Then
                        Set rptRcd = rptRow.Record.Childs.Add()
                    End If
                Next
            End If
            If "" & !性质 = "1" Then
                rptRcd.AddItem("1").Icon = 3
            Else
                rptRcd.AddItem(CStr("" & !性质)).Icon = 2
                mblnStart = True
            End If
            rptRcd.AddItem CStr(!ID)
            rptRcd.AddItem CStr(!判断)
            rptRcd.AddItem CStr("" & !规则)
            rptRcd.AddItem CStr("" & !批范围)
            rptRcd.AddItem CStr("" & !多水平)
            rptRcd.AddItem CStr("" & !符合处理)
            rptRcd.AddItem CStr("" & !不符处理)
            rptRcd.AddItem CStr("" & !是否使用)
            rptRcd.Expanded = True
            .MoveNext
        Loop
    End With
    Me.rptRule.Populate
    
    If lngRuleId <> 0 Then
        For Each rptRow In Me.rptRule.Rows
            If Val(rptRow.Record(mColR.ID).Value) = lngRuleId Then
                Set Me.rptRule.FocusedRow = rptRow
                mlngRuleId = Val(Me.rptRule.FocusedRow.Record(mColR.ID).Value)
                Exit For
            End If
        Next
    End If
    If Me.rptRule.FocusedRow Is Nothing Then
        If Me.rptRule.Rows.Count > 0 Then
            Set Me.rptRule.FocusedRow = Me.rptRule.Rows(0)
            mlngRuleId = Val(Me.rptRule.FocusedRow.Record(mColR.ID).Value)
        Else
            mlngRuleId = 0: Call rptRule_SelectionChanged
        End If
    End If
    zlRefRule = Me.rptRule.Records.Count
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefRule = Me.rptRule.Records.Count
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    If Me.rptRule.Records.Count = 0 Then Exit Sub
    '-------------------------------------------------
    '复制数据表格
    If zlControl.RPTCopyToVSF(Me.rptRule, Me.vfgList) Is Nothing Then Exit Sub
    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "仪器质控规则"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("设备:" & Me.rptDev.FocusedRow.Record(mColD.名称).Value)
    Call objPrint.UnderAppRows.Add(objAppRow)
    
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

Private Sub SetStat()
'
If mintEditState = 0 Then
    '查看
     Me.picDev.Enabled = True:  Me.picRes.Enabled = True: Me.rptRule.SetFocus
ElseIf mintEditState = 1 Then
    '编辑
     Me.picDev.Enabled = False:  Me.picRes.Enabled = False
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

        lngRetuId = mfrmEdit.zlEditSave()
        If lngRetuId <> 0 Then
            mlngRuleId = lngRetuId: Call zlRefRule(mlngRuleId)
            mintEditState = 0: Call SetStat
        End If
        
    Case conMenu_Edit_Untread:
        mfrmEdit.zlEditCancel: Call zlRefRule(mlngRuleId)
        mintEditState = 0: Call SetStat
        
    Case conMenu_Edit_Import
        If frmAppRuleImport.ShowMe(Me, mlngDevId, mlngItemID) Then Call zlRefRule
    Case conMenu_Edit_Archive
        Call frmAppRuleSaveTo.ShowMe(Me, mlngDevId, mlngItemID)
    Case conMenu_Edit_ApplyTo
        '复制项目规则
        If mlngDevId = 0 Then Exit Sub
        If Me.rptDev.FocusedRow Is Nothing Then Exit Sub
        If mlngItemID = 0 Then Exit Sub
        Call frmAppRuleCopy.ShowMe(mlngDevId, mlngItemID, Me)
    Case conMenu_Edit_NewItem
        If mlngDevId = 0 Then Exit Sub
        If Me.rptRule.FocusedRow Is Nothing Then
            If mfrmEdit.zlEditStart(True, 0, mlngDevId, mlngItemID, False) Then
                mintEditState = 1: Call SetStat
            Else
                Call zlRefRule(mlngRuleId)
            End If
        Else
            If Me.rptRule.FocusedRow.Record(mColR.性质).Value = "1" Then
                If mfrmEdit.zlEditStart(True, 0, mlngDevId, mlngItemID, False) Then
                    mintEditState = 1: Call SetStat
                Else
                    Call zlRefRule(mlngRuleId)
                End If
            Else
                If mfrmEdit.zlEditStart(True, mlngRuleId, mlngDevId, mlngItemID, False) Then
                    mintEditState = 1: Call SetStat
                Else
                    Call zlRefRule(mlngRuleId)
                End If
            End If
        End If
        Me.dkpMan.FindPane(conPane_Edit).Select
    
    Case conMenu_Edit_Append
        If mlngDevId = 0 Then Exit Sub
         
        If mfrmEdit.zlEditStart(True, 0, mlngDevId, mlngItemID, True) Then
            mintEditState = 1: Call SetStat
        End If
        Me.dkpMan.FindPane(conPane_Edit).Select
    
    Case conMenu_Edit_Modify
        If mlngDevId = 0 Then Exit Sub
        If mlngRuleId = 0 Then Exit Sub
        If mfrmEdit.zlEditStart(False, mlngRuleId, mlngDevId, mlngItemID) Then
            mintEditState = 1: Call SetStat
        End If
        Me.dkpMan.FindPane(conPane_Edit).Select
    Case conMenu_Edit_Delete
        Dim strMsg As String
        If mlngRuleId = 0 Then Exit Sub
        With Me.rptRule
            strMsg = "真的删除该规则判断吗？"
            strMsg = strMsg & vbCrLf & vbCrLf & "——“" & .FocusedRow.Record(mColR.判断).Value & "”"
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSql = "Zl_检验仪器规则_Edit(3," & mlngRuleId & "," & mlngItemID & ")"
                Err = 0: On Error GoTo ErrHand
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
                
                Err = 0: On Error GoTo 0
                mlngRuleId = 0: lngRetuId = .FocusedRow.Index
                If .Rows.Count > lngRetuId + 1 Then
                    mlngRuleId = .Rows(lngRetuId + 1).Record(mColR.ID).Value
                ElseIf lngRetuId > 0 Then
                    mlngRuleId = .Rows(lngRetuId - 1).Record(mColR.ID).Value
                End If
                Call zlRefRule(mlngRuleId)
            End If
        End With
        Exit Sub
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
    Case conMenu_View_Refresh
        Call zlRefRule(mlngRuleId)
    
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
    Dim intChilds As Integer
    
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptRule.Records.Count > 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (mintEditState <> 0)
    Case conMenu_Edit_Import
        Control.Enabled = (InStr(1, mstrPrivs, "设置") > 0 And mintEditState = 0) And mlngDevId <> 0
    Case conMenu_Edit_Archive
        Control.Enabled = (InStr(1, mstrPrivs, "范例") > 0 And mintEditState = 0) And mlngDevId <> 0 And mblnStart = True
    
    Case conMenu_Edit_NewItem
        Control.Enabled = (InStr(1, mstrPrivs, "设置") > 0 And mintEditState = 0) And mlngDevId <> 0
        If Control.Enabled = False Then Exit Sub
        If Me.rptRule.FocusedRow Is Nothing Then
            Control.Enabled = (mblnStart = False)
        Else
            If Me.rptRule.FocusedRow.Record(mColR.性质).Value = "1" Then
                Control.Enabled = (mblnStart = False)
            Else
                intChilds = 0
                If Me.rptRule.FocusedRow.Record(mColR.符合处理).Value <> "结束" Then intChilds = intChilds + 1
                If Me.rptRule.FocusedRow.Record(mColR.不符处理).Value <> "结束" Then intChilds = intChilds + 1
                Control.Enabled = (Me.rptRule.FocusedRow.Childs.Count < intChilds)
            End If
        End If
    Case conMenu_Edit_Append
        Control.Enabled = (InStr(1, mstrPrivs, "设置") > 0 And mintEditState = 0) And mlngDevId <> 0
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_ApplyTo
        Control.Enabled = (InStr(1, mstrPrivs, "设置") > 0 And mintEditState = 0 And mlngRuleId <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible

    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Dev
        Item.Handle = Me.picDev.hWnd
    Case conPane_Res
        Item.Handle = Me.picRes.hWnd
    Case conPane_Edit
        If mfrmEdit Is Nothing Then Set mfrmEdit = New frmAppRuleEdit
        Item.Handle = mfrmEdit.hWnd
    End Select
End Sub

Private Sub Form_Activate()
    Call SetStat
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    
    mLngEditWidth = frmAppRuleEdit.Width
    mLngEditHeight = frmAppRuleEdit.Height
    
    mintEditState = 0
    mlngDevId = 0: mlngRuleId = 0
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "从范例导入(&I)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "另存到范例(&A)…")
        cbrControl.Style = xtpButtonCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新多规则项(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "新附加规则(&E)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "复制规则(&C)"): cbrControl.BeginGroup = True
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
        .Add FCONTROL, Asc("E"), conMenu_Edit_Append
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
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
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "从范例导入"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新多规则项"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "复制"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "新附加规则")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panThis As Pane, panSub1 As Pane
    
    If mfrmEdit Is Nothing Then Set mfrmEdit = New frmAppRuleEdit
    
    Set panThis = dkpMan.CreatePane(conPane_Dev, 240, 600, DockLeftOf, Nothing)
    panThis.Title = "检验仪器列表"
    panThis.Options = PaneNoCaption
    Set panThis = dkpMan.CreatePane(conPane_Res, 600, 300, DockRightOf, Nothing)
    panThis.Title = "应用质控规则"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panSub1 = dkpMan.CreatePane(conPane_Edit, 600, 400, DockBottomOf, panThis)
    panSub1.Title = "质控规则信息"
    panSub1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    '设备列表的设置
    With Me.rptDev
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 1024)   '必须在列设置之前设置，才能生效
        Set rptCol = .Columns.Add(mColD.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColD.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mColD.名称, "仪器名称", 130, True): rptCol.Editable = False: rptCol.Groupable = True
        
        Set rptCol = .Columns.Add(mColD.项目id, "项目ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mColD.编码, "编码", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColD.项目, "中文名", 130, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColD.英文名, "英文名", 30, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mColD.质控周期, "质控周期", 55, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColD.水平数, "水平数", 45, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列项目..."
            .NoItemsText = "没有可显示的项目..."
            
            .VerticalGridStyle = xtpGridSolid
        End With
        .PreviewMode = False
        .FocusSubItems = True
        '加入项目
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mColD.名称)
        .GroupsOrder(0).SortAscending = True
        .Columns.Find(mColD.名称).Visible = False
    
    End With
    
    '-----------------------------------------------------
    '规则列表的设置
    With Me.rptRule
        .AutoColumnSizing = True
        Set rptCol = .Columns.Add(mColR.性质, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColR.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mColR.判断, "判断描述", 150, True): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.TreeColumn = True
        Set rptCol = .Columns.Add(mColR.规则, "判断规则", 82, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.批范围, "批范围", 45, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.多水平, "多水平", 45, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColR.符合处理, "符合处理", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.不符处理, "不符处理", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.是否使用, "是否使用", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列项目..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With

    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '数据装入
    Call zlRefDev

End Sub

Private Sub Form_Resize()
    Dim panBase As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Set panBase = Me.dkpMan.FindPane(conPane_Edit)
    panBase.MinTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 0
    panBase.MaxTrackSize.SetSize Screen.Width, mLngEditHeight / Screen.TwipsPerPixelY
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters

'    panBase.MinTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, 0
'    panBase.MaxTrackSize.SetSize Screen.Width, mlngEditHeight / Screen.TwipsPerPixelY
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmEdit
    Set mfrmEdit = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picDev_Resize()
    With Me.rptDev
        .Left = Me.picDev.ScaleLeft: .Width = Me.picDev.ScaleWidth - .Left
        .Top = Me.picDev.ScaleTop: .Height = Me.picDev.ScaleHeight - .Top
    End With
End Sub

Private Sub picRes_Resize()
    With Me.rptRule
        .Left = Me.picRes.ScaleLeft
        .Width = Me.picRes.ScaleWidth - .Left
        
        .Top = Me.picRes.ScaleTop + 10
        .Height = Me.picRes.ScaleHeight - .Top
    End With
End Sub

Private Sub rptDev_SelectionChanged()
    
    If Me.rptDev.FocusedRow Is Nothing Then
        mlngDevId = 0
    ElseIf Me.rptDev.FocusedRow.GroupRow Then
        mlngDevId = 0
    Else
        mlngDevId = Me.rptDev.FocusedRow.Record.Item(mColD.ID).Value
        mlngItemID = Me.rptDev.FocusedRow.Record.Item(mColD.项目id).Value
    End If
    Call zlRefRule
End Sub

Private Sub rptRule_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptRule.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptRule.FocusedRow Is Nothing Then Exit Sub
    If Me.rptRule.FocusedRow.GroupRow Then Exit Sub
    Call rptRule_RowDblClick(Me.rptRule.FocusedRow, Me.rptRule.FocusedRow.Record.Item(mColR.ID))
End Sub

Private Sub rptRule_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
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

Private Sub rptRule_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub rptRule_SelectionChanged()
    If Me.rptRule.FocusedRow Is Nothing Then
        mlngRuleId = 0
    Else
        mlngRuleId = Me.rptRule.FocusedRow.Record.Item(mColD.ID).Value
    End If
    Call mfrmEdit.zlRefresh(mlngRuleId)
End Sub





