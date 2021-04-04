VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContentCopy 
   Caption         =   "专用复制"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   Icon            =   "frmContentCopy.frx":0000
   LinkTopic       =   "专用复制"
   ScaleHeight     =   9825
   ScaleWidth      =   14190
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl RptThis 
      Height          =   4440
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      _Version        =   589884
      _ExtentX        =   5106
      _ExtentY        =   7832
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
   End
   Begin VB.Frame fraThis 
      Height          =   700
      Left            =   5280
      TabIndex        =   1
      Top             =   4200
      Width           =   3135
      Begin VB.CommandButton cmdCancle 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   360
         Left            =   1935
         TabIndex        =   3
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "插入选中内容(&I)"
         Height          =   360
         Left            =   300
         TabIndex        =   2
         Top             =   240
         Width           =   1605
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5760
      Top             =   5400
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
            Picture         =   "frmContentCopy.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":6DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":7386
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":7920
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":7EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":8454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmContentCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1
Private mblnOk As Boolean
Private Enum mCol
    ID = 0: 主页ID: 病人ID: 病历种类: 完成时间: 病历名称: 编辑方式: 病人来源: 入院日期: 创建时间:
End Enum
Public Function ShowMe(ByVal frmParent As Object, ByVal patiantID As String, ByVal patiantPageId As String, ByVal lngPatiFrom As Long) As Boolean
    mblnOk = False
    Call Me.zlRefresh(patiantID, patiantPageId, lngPatiFrom)
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Function CopyEnable() As Integer
On Error GoTo errHand
Dim lngRecordId As Long
    
    On Error GoTo errHand
    If Me.RptThis.FocusedRow Is Nothing Then
        Exit Function
    End If
    If Me.RptThis.FocusedRow.Record Is Nothing Then
        Exit Function
    End If
    lngRecordId = Me.RptThis.FocusedRow.Record.Item(mCol.ID).Value
    
	Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Zl_Fun_CopyEnable([1]) CopyEnable From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    If rsTemp!CopyEnable = 1 Then
        CopyEnable = 1
    Else
        CopyEnable = 0
    End If
    
    Exit Function
errHand:
    CopyEnable = 0
End Function

Private Sub cmdInsert_Click()
    If Not mfrmContent Is Nothing Then
        If mfrmContent.edtThis.SelText <> "" Then
            If CopyEnable() = 1 Then
                mfrmContent.edtThis.Copy    '允许以文本方式拷贝到其他程序（放到剪贴板）
                mblnOk = True
                Unload Me
            Else
                MsgBox "选定的病历不允许复制", vbInformation, gstrSysName
            End If
        Else
            mblnOk = False
            MsgBox "请先选择内容！", vbOKOnly + vbInformation, gstrSysName
        End If
    End If
        
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
Select Case Item.ID
        Case 1
        Item.Handle = Me.RptThis.hWnd
        Case 2
        Item.Handle = mfrmContent.hWnd
        Case 3
        Item.Handle = Me.fraThis.hWnd
End Select
End Sub

Private Sub dkpMan_Resize()
    Me.cmdInsert.Move Me.fraThis.Width - Me.cmdInsert.Width - Me.cmdCancle.Width - 200, 160
    Me.cmdCancle.Move Me.fraThis.Width - Me.cmdCancle.Width - 200, 160
End Sub


Private Sub Form_Load()
    Dim rptCol As ReportColumn
    Dim panList As Pane, panContent As Pane, panNew As Pane
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    Set panList = dkpMan.CreatePane(1, 200, 100, DockLeftOf, Nothing)
    panList.MaxTrackSize.Width = 270
    panList.Title = "病历列表"
    panList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set mfrmContent = New frmDockEPRContent
    mfrmContent.mIsShowAnnex = True
    Set panContent = dkpMan.CreatePane(2, 200, 300, DockRightOf, Nothing)
    panContent.Title = "病历内容"
    panContent.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panNew = dkpMan.CreatePane(3, 100, 40, DockBottomOf, panContent)
    panNew.MaxTrackSize.Height = 40
    panNew.Options = PaneNoFloatable Or PaneNoHideable
    panNew.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    With Me.RptThis
        Set rptCol = .Columns.Add(mCol.ID, "ID", 110, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.主页ID, "主页ID", 110, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.病人ID, "病人ID", 110, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.病历种类, "病历种类", 20, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.完成时间, "完成时间", 120, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.病历名称, "病历名称", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.编辑方式, "编辑方式", 120, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.病人来源, "来源", 120, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.入院日期, "来源", 120, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.创建时间, "创建时间", 120, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
'        '.SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .ShowHeader = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
        If Me.RptThis.Rows.Count > 0 Then
            'Me.RptThis.Rows(1).Selected = True
            'Call mfrmContent.zlRefresh(Me.RptThis.Rows(1).Record(mCol.ID).Value, "NOUSE")
        End If
End Sub
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatiFrom As Long) As Boolean
    '功能：刷新装入指定种类的病历文件清单，并定位到指定的文件上
Dim strGroups As String
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    Me.RptThis.Tag = ""
    Me.RptThis.SetImageList Me.imgList

    gstrSQL = "Select ID, 序号, 病人id, 主页id, 病人来源, 病历名称, 完成时间, 创建时间, 病历种类, 编辑方式, 入院日期" & vbNewLine & "From ("
    
    If lngPatiFrom = 2 Or InStr(gstrPrivsEpr, "历史文件") <> 0 Then
        gstrSQL = gstrSQL & _
                    "       Select r.Id, r.序号, r.病人id, r.主页id, r.病人来源, r.病历名称, To_Char(r.完成时间, 'yyyy-mm-dd hh24:mi') As 完成时间, r.创建时间, r.病历种类," & vbNewLine & _
                    "       r.编辑方式, '第' || LPad(r.主页id, 2, '0') || '次住院病历' || '(' || To_Char(m.入院日期, 'YYYY-MM-DD HH24:MI:SS') || ')' As 入院日期" & vbNewLine & _
                    "       From 电子病历记录 R, 病案主页 M" & vbNewLine & _
                    "       Where r.病历种类 In (2, 5, 6) And nvl(R.编辑方式,0)=0 And m.病人id = r.病人id And m.主页id = r.主页id And r.病人id = [1] And r.病人来源 = 2"
        If InStr(gstrPrivsEpr, "历史文件") = 0 Then '没权限只能看本次就诊
            gstrSQL = gstrSQL & " And r.主页id=[2] "
        End If
        gstrSQL = gstrSQL & "       Union" & vbNewLine & _
                    "       Select r.Id, r.序号, r.病人id, r.主页id, r.病人来源, r.病历名称, To_Char(r.完成时间, 'yyyy-mm-dd hh24:mi') As 完成时间, r.创建时间, r.病历种类," & vbNewLine & _
                    "       r.编辑方式, '第' || LPad(r.主页id, 2, '0') || '次住院病历' || '(' || To_Char(m.入院日期, 'yyyy-mm-dd hh24:mi:ss') || ')' As 入院日期" & vbNewLine & _
                    "       From 电子病历记录 R, 病案主页 M, 病人医嘱报告 L, 病人医嘱记录 A" & vbNewLine & _
                    "       Where r.病历种类 = 7 And nvl(r.编辑方式,0)=0 And r.Id = l.病历id And l.医嘱id = a.Id And a.诊疗类别 In('D','E') And m.病人id = r.病人id And" & vbNewLine & _
                    "             m.主页id = r.主页id And r.病人id = [1] And r.病人来源 = 2" & vbNewLine & _
                   "        And (Exists (Select 1 From 影像检查记录 Where 医嘱ID=a.ID) or l.RISID IS NOT NULL)"
        If InStr(gstrPrivsEpr, "历史文件") = 0 Then '没权限只能看本次就诊
            gstrSQL = gstrSQL & " And r.主页id=[2] "
        End If
        If InStr(GetPrivFunc(glngSys, IIf(lngPatiFrom = 2, 1253, 1252)), "查阅未完成报告") = 0 Then '没权限时不能查看未完成报告
            gstrSQL = gstrSQL & vbNewLine & " And Exists (Select 1 From 病人医嘱发送 E Where E.医嘱ID=A.ID And (E.执行过程>=5 or E.执行状态=1))"
        End If
    End If
    
    If lngPatiFrom = 1 Or InStr(gstrPrivsEpr, "历史文件") <> 0 Then
         If lngPatiFrom = 1 And InStr(gstrPrivsEpr, "历史文件") = 0 Then
         Else
            gstrSQL = gstrSQL & "       Union" & vbNewLine
         End If
         
        gstrSQL = gstrSQL & _
                    "       Select r.id, r.序号, r.病人id, r.主页id, r.病人来源, r.病历名称, To_Char(r.完成时间, 'yyyy-mm-dd hh24:mi') As 完成时间,r.创建时间, r.病历种类," & vbNewLine & _
                    "       r.编辑方式, '门诊病历'||'('||to_char(nvl(m.执行时间,m.登记时间),'yyyy-mm-dd hh24:mi:ss')||')' as 入院日期 " & vbNewLine & _
                    "       From 电子病历记录 r,病人挂号记录  m " & vbNewLine & _
                    "       Where r.病历种类 in (1,5,6) And nvl(r.编辑方式,0)=0 and M.病人ID = r.病人ID and m.ID=r.主页id And r.病人ID = [1] And r.病人来源 = 1"
        If InStr(gstrPrivsEpr, "历史文件") = 0 Then '没权限只能看本次就诊
            gstrSQL = gstrSQL & " And r.主页id=[2] "
        End If
        gstrSQL = gstrSQL & "       Union" & vbNewLine & _
                    "       Select r.id, r.序号, r.病人id, r.主页id, r.病人来源, r.病历名称, To_Char(r.完成时间, 'yyyy-mm-dd hh24:mi') As 完成时间,r.创建时间, r.病历种类," & vbNewLine & _
                    "       r.编辑方式, '门诊病历'||'('||to_char(nvl(m.执行时间,m.登记时间),'yyyy-mm-dd hh24:mi:ss')||')' as 入院日期 " & vbNewLine & _
                    "       From 电子病历记录 r,病人挂号记录  m,病人医嘱报告 L ,病人医嘱记录 A" & vbNewLine & _
                    "       Where r.病历种类  = 7 And nvl(r.编辑方式,0)=0 and M.病人ID = r.病人ID and m.ID=r.主页id And r.Id = l.病历id And l.医嘱id = a.Id And a.诊疗类别 In('D','E') And r.病人ID = [1] And r.病人来源 = 1" & vbNewLine & _
                   "        And (Exists (Select 1 From 影像检查记录 Where 医嘱ID=a.ID) or l.RISID IS NOT NULL)"
        If InStr(gstrPrivsEpr, "历史文件") = 0 Then '没权限只能看本次就诊
            gstrSQL = gstrSQL & " And r.主页id=[2] "
        End If
        If InStr(GetPrivFunc(glngSys, IIf(lngPatiFrom = 2, 1253, 1252)), "查阅未完成报告") = 0 Then '没权限时不能查看未完成报告
            gstrSQL = gstrSQL & vbNewLine & " And Exists (Select 1 From 病人医嘱发送 E Where E.医嘱ID=A.ID And (E.执行过程>=5 or E.执行状态=1))"
        End If
    End If
    gstrSQL = gstrSQL & ") Order By 入院日期 DESC, 创建时间 ASC"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPageId)
    Me.RptThis.Records.DeleteAll
    With rsTemp
        strGroups = ""
        Do While Not .EOF
            Set rptRcd = Me.RptThis.Records.Add()
            rptRcd.AddItem CStr(!ID)
            rptRcd.AddItem CStr(!主页ID)
            rptRcd.AddItem CStr(!病人ID)
            rptRcd.AddItem (CStr(!病历种类))
            rptRcd.AddItem CStr(NVL(!完成时间, ""))
            Set rptItem = rptRcd.AddItem(CStr(!病历名称)): rptItem.Icon = NVL(!病历种类, 0) - 1
            rptRcd.AddItem CStr(!编辑方式)
            rptRcd.AddItem CStr(!病人来源)
            rptRcd.AddItem CStr(NVL(!入院日期, ""))
            rptRcd.AddItem CStr(NVL(!创建时间))
            .MoveNext
        Loop
        With Me.RptThis
            .SortOrder.Add .Columns.Find(mCol.ID)
            .SortOrder.Add .Columns.Find(mCol.创建时间)
            .SortOrder.Column(0).SortAscending = False
            .SortOrder.Column(1).SortAscending = True
            .GroupsOrder.Add .Columns.Find(mCol.入院日期)
            .GroupsOrder(0).SortAscending = False
            .Populate
        End With
    End With
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
        Unload mfrmContent
        Set mfrmContent = Nothing
End Sub

Private Sub RptThis_SelectionChanged()
    Dim lngRecordId As Long
    On Error GoTo errHand
    If Me.RptThis.FocusedRow Is Nothing Then
        Exit Sub
    End If
    If Me.RptThis.FocusedRow.Record Is Nothing Then
        Exit Sub
    End If
    lngRecordId = Me.RptThis.FocusedRow.Record.Item(mCol.ID).Value
    If Val(Me.RptThis.Tag) <> Me.RptThis.FocusedRow.Index Then
        mfrmContent.mIsShowAnnex = False
        Call mfrmContent.zlRefresh(lngRecordId, "NOUSE", , , , , , True)
        RptThis.Tag = Me.RptThis.FocusedRow.Index
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



