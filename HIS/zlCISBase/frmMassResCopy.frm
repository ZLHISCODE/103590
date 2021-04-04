VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmMassResCopy 
   Caption         =   "质控品复制"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11160
   Icon            =   "frmMassResCopy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   11160
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   6030
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   8070
      _Version        =   589884
      _ExtentX        =   14235
      _ExtentY        =   10636
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   5355
      Left            =   8190
      ScaleHeight     =   5355
      ScaleWidth      =   2925
      TabIndex        =   1
      Top             =   15
      Width           =   2925
      Begin VB.Frame fraLine 
         Height          =   15
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   3285
         Width           =   2760
      End
      Begin VB.Frame fraLine 
         Height          =   15
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   2400
         Width           =   2760
      End
      Begin VB.Frame fraLine 
         Height          =   15
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   1020
         Width           =   2760
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "全清(&C)"
         Height          =   350
         Index           =   1
         Left            =   1485
         TabIndex        =   13
         Top             =   1905
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "全选(&A)"
         Height          =   350
         Index           =   0
         Left            =   375
         TabIndex        =   12
         Top             =   1905
         Width           =   1100
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "关闭(&X)"
         Height          =   350
         Left            =   1485
         TabIndex        =   10
         Top             =   4305
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确认复制(&O)"
         Height          =   350
         Left            =   150
         TabIndex        =   9
         Top             =   4305
         Width           =   1320
      End
      Begin VB.OptionButton optValue 
         Caption         =   "按原批号最后控制值(&2)"
         Height          =   180
         Index           =   1
         Left            =   375
         TabIndex        =   8
         Top             =   3975
         Value           =   -1  'True
         Width           =   2310
      End
      Begin VB.OptionButton optValue 
         Caption         =   "按原批号预设控制值(&1)"
         Height          =   180
         Index           =   0
         Left            =   375
         TabIndex        =   7
         Top             =   3720
         Width           =   2310
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "按原批号顺加一年(&D)"
         Height          =   350
         Left            =   375
         TabIndex        =   5
         Top             =   2775
         Width           =   2220
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "仅已选择的质控品(&S)"
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   4
         Top             =   1650
         Width           =   2070
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "仅当前在用质控品(&U)"
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   3
         Top             =   1395
         Value           =   1  'Checked
         Width           =   2070
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "新批号使用日期范围:"
         Height          =   180
         Left            =   165
         TabIndex        =   17
         Top             =   2565
         Width           =   1710
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         Caption         =   "列表显示范围与选择:"
         Height          =   180
         Left            =   165
         TabIndex        =   11
         Top             =   1140
         Width           =   1710
      End
      Begin VB.Label lblValues 
         AutoSize        =   -1  'True
         Caption         =   "定值新批号的预设控制值:"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   3450
         Width           =   2070
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "    需要批量更换新批号质控品时使用本功能，可快速建立新批号控制品并继承原批号的检测项目和控制值。"
         Height          =   720
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   2700
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   165
         Picture         =   "frmMassResCopy.frx":058A
         Top             =   60
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   270
      Top             =   5385
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMassResCopy.frx":0B14
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMassResCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    选择 = 0: 仪器: ID: 名称: 原批号: 原开始日期: 原结束日期: 新批号: 新开始日期: 新结束日期
End Enum

'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long

'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Public Function zlRefList() As Long
    '功能：刷新装入清单
    Dim rsTemp As New ADODB.Recordset
    gstrSql = "Select R.ID, R.名称, R.批号 As 原批号, R.开始日期 As 原开始日期, R.结束日期 As 原结束日期," & vbNewLine & _
            "       D.编码 || '-' || D.名称 As 仪器" & vbNewLine & _
            "From 检验质控品 R, 检验仪器 D" & vbNewLine & _
            "Where R.仪器id = D.ID"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr("")): rptItem.HasCheckbox = True: rptItem.Checked = False
            rptRcd.AddItem CStr("" & !仪器)
            rptRcd.AddItem CStr("" & !ID)
            rptRcd.AddItem CStr("" & !名称)
            rptRcd.AddItem CStr("" & !原批号)
            rptRcd.AddItem CStr("" & !原开始日期)
            rptRcd.AddItem CStr("" & !原结束日期)
            rptRcd.AddItem CStr("")
            rptRcd.AddItem CStr("")
            rptRcd.AddItem CStr("")
            .MoveNext
        Loop
    End With
    Me.rptList.Populate
    
    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    
    zlRefList = Me.rptList.Records.Count
    Call chkShow_Click(0)
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------

Private Sub chkShow_Click(Index As Integer)
    Dim strNow As String
    Dim blnShow As Boolean
    
    strNow = Format(Now(), "yyyy-MM-dd")
    
    For Each rptRcd In Me.rptList.Records
        blnShow = True
        If Me.chkShow(0).Value = vbChecked Then
            If strNow < rptRcd.Item(mCol.原开始日期).Value Or strNow > rptRcd.Item(mCol.原结束日期).Value Then blnShow = False
        End If
        If Me.chkShow(1).Value = vbChecked Then
            If rptRcd.Item(mCol.选择).Checked = False Then blnShow = False
        End If
        rptRcd.Visible = blnShow
    Next
    Me.rptList.Populate
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDate_Click()
    For Each rptRcd In Me.rptList.Records
        rptRcd.Item(mCol.新开始日期).Value = Format(DateAdd("m", 12, rptRcd.Item(mCol.原开始日期).Value), "yyyy-MM-dd")
        rptRcd.Item(mCol.新结束日期).Value = Format(DateAdd("m", 12, rptRcd.Item(mCol.原结束日期).Value), "yyyy-MM-dd")
    Next
    Me.rptList.Populate
End Sub

Private Sub cmdOK_Click()
    Dim strCopies As String
    strCopies = ""
    For Each rptRcd In Me.rptList.Records
        If rptRcd.Visible And rptRcd.Item(mCol.选择).Checked Then
            If Trim(rptRcd.Item(mCol.新批号).Value) = "" Then
                MsgBox "选择复制的质控品：" & vbCrLf & rptRcd.Item(mCol.名称).Value & " 未设置新批号！", vbInformation, gstrSysName
                Set Me.rptList.FocusedRow = rptRcd: Exit Sub
            End If
            If Trim(rptRcd.Item(mCol.新开始日期).Value) = "" Then
                MsgBox "选择复制的质控品：" & vbCrLf & rptRcd.Item(mCol.名称).Value & " 未设置新开始日期！", vbInformation, gstrSysName
                Set Me.rptList.FocusedRow = rptRcd: Exit Sub
            End If
            If Trim(rptRcd.Item(mCol.新结束日期).Value) = "" Then
                MsgBox "选择复制的质控品：" & vbCrLf & rptRcd.Item(mCol.名称).Value & " 未设置新结束日期！", vbInformation, gstrSysName
                Set Me.rptList.FocusedRow = rptRcd: Exit Sub
            End If
            strCopies = strCopies & "|" & rptRcd.Item(mCol.ID).Value
            strCopies = strCopies & ";" & rptRcd.Item(mCol.新批号).Value
            strCopies = strCopies & ";" & rptRcd.Item(mCol.新开始日期).Value
            strCopies = strCopies & ";" & rptRcd.Item(mCol.新结束日期).Value
        End If
    Next
    If strCopies = "" Then MsgBox "尚未选择复制的质控品！", vbInformation, gstrSysName: Exit Sub
    gstrSql = "Zl_检验质控品_Copy(" & IIf(Me.optValue(0).Value, 0, 1) & ",'" & Mid(strCopies, 2) & "')"
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    For Each rptRcd In Me.rptList.Records
        rptRcd.Item(mCol.选择).HasCheckbox = True
        rptRcd.Item(mCol.选择).Checked = (Index = 0)
    Next
    Call chkShow_Click(0)
End Sub

Private Sub Form_Load()
    With Me.rptList
        .AutoColumnSizing = True
        Set rptCol = .Columns.Add(mCol.选择, "", 18, False): rptCol.Editable = True: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.仪器, "仪器", 120, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.名称, "名称", 150, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.原批号, "原批号", 66, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.原开始日期, "原开始日期", 66, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.原结束日期, "原结束日期", 66, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.新批号, "新批号", 60, False): rptCol.Editable = True: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.新开始日期, "新开始日期", 66, False): rptCol.Editable = True: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.新结束日期, "新结束日期", 66, False): rptCol.Editable = True: rptCol.Groupable = False
        
        .AllowEdit = True
        .EditOnClick = True
        .FocusSubItems = True
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
        .GroupsOrder.Add .Columns.Find(mCol.仪器)
        .GroupsOrder(0).SortAscending = True
    End With
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)

    Call zlRefList
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.picBack
        .Left = Me.ScaleWidth - .Width
        .Height = Me.ScaleHeight
    End With
    With Me.rptList
        .Left = 0: .Width = Me.picBack.Left
        .Top = 0: .Height = Me.ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub rptList_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim dtInput As Date
    If Trim(Item.Value) = "" Then Item.Value = CStr(""): Exit Sub
    Select Case Column.Index
    Case mCol.新批号
        Item.Value = UCase(Item.Value)
        For lngCount = 1 To Len(Item.Value)
            If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(Item.Value, lngCount, 1)) = 0 Then
                MsgBox "批号包含不允许的字符", vbInformation, gstrSysName
                Item.Value = CStr(""): Exit Sub
            End If
        Next
        If Len(Item.Value) > 10 Then
            MsgBox "批号太长！", vbInformation, gstrSysName
            Item.Value = CStr(""): Exit Sub
        End If
        If Item.Value = Row.Record.Item(mCol.原批号).Value Then
            MsgBox "新批号不能和原批号相同！", vbInformation, gstrSysName
            Item.Value = CStr(""): Exit Sub
        End If
    Case mCol.新开始日期, mCol.新结束日期
        Err = 0: On Error Resume Next
        dtInput = CDate(Item.Value)
        If Err <> 0 Then
            MsgBox "不符合日期格式！", vbInformation, gstrSysName
            Item.Value = CStr(""): Exit Sub
        End If
        
        Err = 0: On Error GoTo 0
        Item.Value = Format(dtInput, "yyyy-MM-dd")
        
        If Row.Record.Item(mCol.新开始日期).Value <> CStr("") And Row.Record.Item(mCol.新结束日期).Value <> CStr("") Then
            If Row.Record.Item(mCol.新开始日期).Value >= Row.Record.Item(mCol.新结束日期).Value Then
                MsgBox "新的日期范围错误(结束日期必须大于开始日期)！", vbInformation, gstrSysName
                Item.Value = CStr(""): Exit Sub
            End If
        End If
    End Select
End Sub
