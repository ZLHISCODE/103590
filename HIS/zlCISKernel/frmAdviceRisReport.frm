VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAdviceRisReport 
   AutoRedraw      =   -1  'True
   Caption         =   "打印RIS预约单"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   13260
   Icon            =   "frmAdviceRisReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   13260
   StartUpPosition =   2  '屏幕中心
   Begin XtremeReportControl.ReportControl rptAdvice 
      Height          =   1170
      Left            =   2205
      TabIndex        =   21
      Top             =   2415
      Width           =   630
      _Version        =   589884
      _ExtentX        =   1111
      _ExtentY        =   2064
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.TabControl tbcAppend 
      Height          =   1530
      Left            =   3015
      TabIndex        =   24
      Top             =   5325
      Width           =   270
      _Version        =   589884
      _ExtentX        =   476
      _ExtentY        =   2699
      _StockProps     =   64
   End
   Begin VB.Frame fraAdviceUD 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   -405
      MousePointer    =   7  'Size N S
      TabIndex        =   22
      Top             =   4725
      Width           =   6000
   End
   Begin VB.PictureBox picDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   3855
      ScaleHeight     =   2850
      ScaleWidth      =   4890
      TabIndex        =   14
      Top             =   1980
      Visible         =   0   'False
      Width           =   4920
      Begin VB.CommandButton cmdFindCancle 
         Caption         =   "取消"
         Height          =   270
         Left            =   4200
         TabIndex        =   7
         Top             =   75
         Width           =   615
      End
      Begin VB.CommandButton cmdFindOk 
         Caption         =   "确定"
         Height          =   270
         Left            =   3480
         TabIndex        =   8
         Top             =   75
         Width           =   615
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找"
         Height          =   270
         Left            =   1740
         TabIndex        =   16
         Top             =   75
         Width           =   615
      End
      Begin VB.TextBox txtFind 
         Height          =   270
         Left            =   50
         TabIndex        =   15
         Top             =   75
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2280
         Left            =   75
         TabIndex        =   9
         ToolTipText     =   "全选Ctrl+A；全清Ctrl+R"
         Top             =   510
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   4022
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img16"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame fraFilter 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   375
      Width           =   12300
      Begin VB.CommandButton cmdFilter 
         Caption         =   "查找(F3)"
         Height          =   300
         Left            =   4230
         TabIndex        =   6
         Top             =   765
         Width           =   900
      End
      Begin VB.TextBox txtFilter 
         Height          =   300
         Left            =   2205
         TabIndex        =   5
         Top             =   780
         Width           =   2000
      End
      Begin VB.ComboBox cboFind 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   780
         Width           =   1320
      End
      Begin VB.OptionButton optType 
         Caption         =   "已打印"
         Height          =   240
         Index           =   1
         Left            =   11130
         TabIndex        =   19
         Top             =   405
         Width           =   870
      End
      Begin VB.OptionButton optType 
         Caption         =   "未打印"
         Height          =   240
         Index           =   0
         Left            =   10260
         TabIndex        =   18
         Top             =   405
         Value           =   -1  'True
         Width           =   850
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   0
         Width           =   2160
      End
      Begin VB.CommandButton cmdDept 
         Caption         =   "…"
         Height          =   265
         Left            =   4560
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl+D"
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox txtDept 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   870
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Text            =   "所有科室"
         ToolTipText     =   "所有科室"
         Top             =   345
         Width           =   4000
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   5820
         TabIndex        =   2
         Top             =   360
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   179765251
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   8130
         TabIndex        =   3
         Top             =   360
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   179765251
         CurrentDate     =   37953
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "查找病人"
         Height          =   180
         Left            =   90
         TabIndex        =   25
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblTim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预约时间                        至"
         Height          =   180
         Left            =   5025
         TabIndex        =   20
         Top             =   405
         Width           =   3060
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "开单科室"
         Height          =   180
         Left            =   90
         TabIndex        =   11
         Top             =   375
         Width           =   720
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "住院病区"
         Height          =   180
         Left            =   90
         TabIndex        =   13
         Top             =   45
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   7425
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   635
      SimpleText      =   $"frmAdviceRisReport.frx":014A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceRisReport.frx":0191
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18309
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
   Begin RichTextLib.RichTextBox rtfAppend 
      Height          =   1395
      Left            =   4320
      TabIndex        =   23
      Top             =   5295
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   2461
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAdviceRisReport.frx":0A25
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
   Begin MSComctlLib.ImageList img16 
      Left            =   660
      Top             =   1875
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":0AC2
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":105C
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":15F6
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":1B90
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":212A
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":26C4
            Key             =   "单病种"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":8F26
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":F788
            Key             =   "printer"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   270
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAdviceRisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng病区ID As Long
Private mlngFind As Long
Private mstrMatch As String
Private mstrPrivs As String
Private mintPreDept As Integer
Private mstrFindType As String

Private Enum PatiCol
    COL_选择
    COL_姓名
    COL_住院号
    COL_床号
    COL_性别
    COL_年龄
    COL_科室
    COL_预约时间
    COL_预约设备
    COL_预约号
    COL_预约项目
    COL_打印时间
    COL_打印人
    
    COL_病人ID
    COL_主页ID
    COL_医嘱ID
End Enum

Private Enum Ectrl
    e未打
    e已打
End Enum

Public Function ShowMe(frmParent As Object, ByVal lng病区ID As Long) As Boolean
    mlng病区ID = lng病区ID
    Me.Show , frmParent
End Function

Private Sub cboFind_Click()
    mstrFindType = cboFind.Text
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngLW As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    fraFilter.Top = lngTop
    fraFilter.Left = lngLeft
    fraFilter.Width = Me.ScaleWidth
    
    rptAdvice.Left = lngLeft
    rptAdvice.Top = fraFilter.Top + fraFilter.Height
    rptAdvice.Width = lngRight - lngLeft
    rptAdvice.Height = lngBottom - rptAdvice.Top - fraAdviceUD.Height - tbcAppend.Height
    
    fraAdviceUD.Left = lngLeft
    fraAdviceUD.Top = rptAdvice.Top + rptAdvice.Height
    fraAdviceUD.Width = rptAdvice.Width
    
    tbcAppend.Left = lngLeft
    tbcAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    tbcAppend.Width = rptAdvice.Width
    Me.Refresh
End Sub

Private Sub PrintRIS()
    Dim i As Long, j As Long
    Dim lngResult As Long
    Dim lng医嘱ID As Long

    If HaveRIS Then
        '病人
        For i = 0 To rptAdvice.Rows.Count - 1
            If Not rptAdvice.Rows(i).GroupRow Then
                If rptAdvice.Rows(i).Record.Tag = "1" Then
                    lng医嘱ID = Val(rptAdvice.Rows(i).Record(COL_医嘱ID).value)
                    lngResult = -1
                    lngResult = gobjRis.HISPrintOneRisScheduleRpt(lng医嘱ID)
                    j = j + 1
                End If
            End If
        Next
        If j = 0 Then
            MsgBox "未勾选任何项目。", vbInformation, gstrSysName
            rptAdvice.SetFocus: Exit Sub
        End If
    End If
End Sub

Private Sub cmdFilter_Click()
    Call ExecuteFindPati
End Sub

Private Sub dtpBegin_KeyPress(KeyAscii As Integer)
    If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEnd_KeyPress(KeyAscii As Integer)
    If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()

    Dim datCur As Date
 
    On Error GoTo errH
    
    mstrPrivs = gMainPrivs
    
    Call InitCommandBar
    
    With tbcAppend
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
        End With
        .InsertItem(0, "申请附项", rtfAppend.hwnd, 0).Tag = "附项"
    End With
    
    Call InitReportColumn
    
    mstrMatch = IIF(Val(zlDatabase.GetPara("输入匹配", , , True)) = 0, "%", "")
    
    datCur = zlDatabase.Currentdate
    
    dtpBegin.value = Format(datCur - 1, "yyyy-MM-dd 00:00:00")
    dtpEnd.value = Format(datCur + 1, "yyyy-MM-dd 23:59:59")
    With cboFind
        .Clear
        .AddItem "姓名"
        .AddItem "床号"
        .AddItem "住院号"
        .ListIndex = 0
    End With
    mstrFindType = "姓名"
    Call InitUnits
    Call LoadDept
    Call LoadAdvice
    mintPreDept = -1
    Call RestoreWinState(Me, App.ProductName)
    Me.WindowState = vbMaximized
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub rptAdvice_SelectionChanged()
'功能：显示附项
    Dim lng医嘱ID As Long
    If rptAdvice.SelectedRows.Count = 0 Then Exit Sub
    With rptAdvice.SelectedRows(0)
        lng医嘱ID = Val(.Record(COL_医嘱ID).value)
    End With
    Call ShowAppend(lng医嘱ID)
End Sub

Private Sub LoadDept()
'功能：科室选择器
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim objItem As ListItem
    
    On Error GoTo errH
    
    txtDept.Text = "所有科室"
    txtDept.ToolTipText = "所有科室"
    txtDept.Tag = ""
    picDept.Visible = False
    txtFind.Text = ""
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1500
        .Add , "编码", "编码", 900
    End With
    
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
        .Width = 3000
    End With
    
    strSql = "select distinct ID,编码,名称" & _
        " from 部门表 D,部门性质说明 T,病区科室对应 a" & _
        " where D.ID=T.部门ID and t.工作性质=[1] and d.id=a.科室id and a.病区id=[2]" & _
        " and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
        " order by d.编码"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "临床", mlng病区ID)
    
    Me.lvwItems.ListItems.Clear
    
    Me.lvwItems.Checkboxes = True
   
    Do Until rsTmp.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称)
        objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
        objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = rsTmp!编码
        objItem.Checked = False
        rsTmp.MoveNext
    Loop
    
    '没有时退出
    If Me.lvwItems.ListItems.Count = 0 Then Exit Sub
    
    Me.lvwItems.ListItems(1).Selected = True
    
    Exit Sub
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mlngFind = Item.Index + 1
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then '全选 Ctrl+A
        Call SetSelect(lvwItems, True)
    End If
    
    If KeyCode = vbKeyR And Shift = vbCtrlMask Then '全消 Ctrl+R
        Call SetSelect(lvwItems, False)
    End If
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim i As Long
    Dim blnIsFind As Boolean
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    For i = mlngFind To lvwItems.ListItems.Count
        If zlCommFun.SpellCode(Mid(lvwItems.ListItems(i).Text, InStr(lvwItems.ListItems(i).Text, "-") + 1)) Like UCase(IIF(mstrMatch <> "", "*", "") & strFind & "*") Or _
                UCase(lvwItems.ListItems(i).Text) Like UCase(IIF(mstrMatch <> "", "*", "") & strFind & "*") Then
            lvwItems.ListItems(i).Selected = True
            lvwItems.ListItems(i).EnsureVisible
            blnIsFind = True
            mlngFind = i + 1
            Exit For
        End If
    Next
    If blnIsFind = False Then
        If mlngFind = 1 Then
            MsgBox "没有找到您查找的科室。", vbInformation, Me.Caption
        Else
            MsgBox "已经是最后一个科室了。", vbInformation, Me.Caption
            mlngFind = 1
        End If
    End If
End Sub

Private Sub cmdFindCancle_Click()
    Call lvwItems_KeyPress(vbKeyEscape)
End Sub

Private Sub cmdFindOk_Click()
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_LostFocus()
    Call picDept_LostFocus
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        If lvwItems.SelectedItem.Checked = False And KeyAscii = vbKeyReturn Then
            lvwItems.SelectedItem.Checked = Not lvwItems.SelectedItem.Checked
            Exit Sub
        End If
        If lvwItems.Checkboxes = True And KeyAscii = vbKeySpace Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
        picDept.Visible = False
        txtFind.Text = ""
    End Select
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal blnSelect As Boolean = True)
    Dim i As Integer
    
    With lvwObj
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next
    End With
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    Dim m As Integer
    Dim blnBatch As Boolean
    Dim str科室 As String
    Dim str科室IDs As String
    Dim strTmp As String
    Dim varArr As Variant
    Dim n As Integer
    Dim strNew As String
    Dim blnNew As Boolean
        
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
  
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Checked Then
            strTmp = Mid(lvwItems.ListItems(i).Key, 2) & "," & lvwItems.ListItems(i).Text
            If InStr(str科室, strTmp) = 0 Then str科室 = str科室 & ";" & strTmp
        End If
    Next
    If str科室 = "" Then
        txtDept.Text = "所有科室"
        txtDept.ToolTipText = "所有科室"
        txtDept.Tag = ""
        picDept.Visible = False
        txtFind.Text = ""
        Exit Sub
    End If
    str科室 = Mid(str科室, 2)
    
    varArr = Split(str科室, ";"): strTmp = ""
    
    For i = 0 To UBound(varArr)
        strTmp = strTmp & "," & Split(varArr(i), ",")(1)
        str科室IDs = str科室IDs & "," & Split(varArr(i), ",")(0)
    Next
    
    txtDept.Text = Mid(strTmp, 2)
    txtDept.ToolTipText = txtDept.Text
    txtDept.Tag = Mid(str科室IDs, 2)
    picDept.Visible = False
    txtFind.Text = ""
End Sub

Private Sub picDept_LostFocus()
    Dim strActive As String
    
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDFINDCANCLE,LVWITEMS,PICDEPT,TXTFIND,CMDFIND,CMDFINDOK", strActive) <> 0 Then
        Exit Sub
    End If

    picDept.Visible = False
    txtFind.Text = ""
    mlngFind = 1
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIF(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub cmdDept_Click()
'功能：显示部门选择器
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim lngTmp  As Long
    Dim i As Integer
    
    With Me.picDept
        .Left = txtDept.Left
        .Width = txtDept.Width + 700
        .Top = txtDept.Top + txtDept.Height + fraFilter.Top
        cmdFind.Visible = True
        txtFind.Visible = True
        cmdFindOk.Visible = True
        cmdFindCancle.Visible = True
        .ZOrder 0
        .Visible = True
    End With

    With Me.lvwItems
        .Left = 0
        .Top = txtFind.Height + 100
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height - txtFind.Height - 50 - 50
        txtFind.Top = 50
        cmdFind.Top = 50
        cmdFindOk.Left = .Width + .Left - cmdFind.Width - 80 - cmdFindCancle.Width
        cmdFindCancle.Left = .Width + .Left - cmdFind.Width - 50
        cmdFindOk.Top = cmdFind.Top
        cmdFindCancle.Top = cmdFind.Top
        .SetFocus
        .Refresh
    End With
    
    Call SetSelect(lvwItems, False)
    If txtDept.Tag = "" Then Exit Sub
   
    For i = 1 To lvwItems.ListItems.Count
        lngTmp = Val(Mid(lvwItems.ListItems(i).Key, 2))
        Me.lvwItems.ListItems(i).Checked = InStr("," & txtDept.Tag & ",", "," & lngTmp & ",") > 0
    Next
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, " 查询"): objControl.BeginGroup = True
        objControl.ToolTipText = "读取RIS检查申请的数据"
            
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "全选")
        objControl.BeginGroup = True
        objControl.ToolTipText = "选中所有可以打印检查申请(Ctrl+A)"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "全清")
        objControl.ToolTipText = "清除所有已选择检查申请的选择状态(Ctrl+R)"
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        objControl.ToolTipText = "对已经勾选的检查申请单执行打印的操作"
        objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " 退出"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print
        .Add 0, vbKeyF5, conMenu_View_Refresh
        .Add FALT, vbKeyX, conMenu_File_Exit
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Refresh
        Call LoadAdvice
    Case conMenu_Edit_SelAll
        Call SelAllCls(True)
    Case conMenu_Edit_ClsAll
        Call SelAllCls(False)
    Case conMenu_File_Print
        Call PrintRIS
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not picDept.Visible Then
        If KeyCode = vbKeyA And Shift = vbCtrlMask Then
            cbsMain.FindControl(, conMenu_Edit_SelAll).Execute
        ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
            cbsMain.FindControl(, conMenu_Edit_ClsAll).Execute
        End If
        If KeyCode = vbKeyF3 Then
            If txtFilter.Text = "" Then
                txtFilter.SetFocus
            Else
                Call ExecuteFindPati(True)
            End If
        End If
    End If
End Sub

Private Sub fraAdviceUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If rptAdvice.Height + Y < 1000 Or tbcAppend.Height - Y < 500 Then Exit Sub
        fraAdviceUD.Top = fraAdviceUD.Top + Y
        rptAdvice.Height = rptAdvice.Height + Y
        tbcAppend.Top = tbcAppend.Top + Y
        tbcAppend.Height = tbcAppend.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn
    
    With rptAdvice
        Set objCol = .Columns.Add(COL_选择, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("UnCheck").Index - 1
        Set objCol = .Columns.Add(COL_姓名, "姓名", 120, True)
        Set objCol = .Columns.Add(COL_住院号, "住院号", 100, True)
        Set objCol = .Columns.Add(COL_床号, "床号", 45, True)
        Set objCol = .Columns.Add(COL_性别, "性别", 30, True)
        Set objCol = .Columns.Add(COL_年龄, "年龄", 45, True)
        Set objCol = .Columns.Add(COL_科室, "科室", 120, True)
        Set objCol = .Columns.Add(COL_预约时间, "预约时间", 120, True)
        Set objCol = .Columns.Add(COL_预约设备, "预约设备", 120, True)
        Set objCol = .Columns.Add(COL_预约号, "预约号", 60, True)
        Set objCol = .Columns.Add(COL_预约项目, "预约项目", 120, True)
	Set objCol = .Columns.Add(COL_打印时间, "打印时间", 120, True)
        Set objCol = .Columns.Add(COL_打印人, "打印人", 120, True)
        
        Set objCol = .Columns.Add(COL_病人ID, "病人ID", 0, False)
        Set objCol = .Columns.Add(COL_主页ID, "主页ID", 0, False)
        Set objCol = .Columns.Add(COL_医嘱ID, "医嘱ID", 0, False)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
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
        .SetImageList Me.img16
    End With
End Sub

Private Sub LoadAdvice()
'功能：加载病人医嘱列表
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strDepts  As String
    Dim strTmp As String
    
    On Error GoTo errH
    
    If dtpBegin.value >= dtpEnd.value Then
        MsgBox "开始时间应小于结束时间。", vbInformation, gstrSysName
        dtpBegin.SetFocus: Exit Sub
    End If
    
    strDepts = txtDept.Tag
    
    strSql = "select b.姓名,b.住院号,b.出院病床 As 床号,nvl(c.婴儿性别,b.性别) as 性别,b.年龄,e.名称 as 科室,f.预约开始时间 as 预约时间,f.检查设备名称 as 预约设备,f.序号 as 预约号,a.医嘱内容 as 预约项目," & vbNewLine & _
        "a.id as 医嘱ID,a.病人id,a.主页id,c.序号 as 婴儿,c.婴儿姓名,Round(Decode(c.死亡时间, Null, Sysdate, c.死亡时间) - c.出生时间)||'天' As 婴儿年龄,b.病人类型,to_char(f.打印时间,'YYYY-MM-DD HH24:MI') as 打印时间,f.打印人" & vbNewLine & _
        "from 病人医嘱记录 a,病案主页 b,病人新生儿记录 c,部门表 e,RIS检查预约 f,病人医嘱发送 g" & vbNewLine & _
        "where a.病人id=b.病人id and a.主页id=b.主页id and a.病人id=c.病人id(+) and a.主页id=c.主页id(+) and a.婴儿=c.序号(+) and a.id=g.医嘱id and nvl(g.执行过程,0)<3 " & vbNewLine & _
        "and a.开嘱科室id=e.id and a.id=f.医嘱id And b.当前病区id =[1] And f.预约日期 between [2] and [3] And nvl(f.是否打印,0)=[4]" & _
        IIF(strDepts = "", "", "  and a.开嘱科室id in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([5]) As zlTools.t_Numlist)) X)") & _
        " order by a.病人id,a.主页id,a.婴儿,f.序号"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病区ID, CDate(dtpBegin.value), CDate(dtpEnd.value), IIF(optType(e未打).value, 0, 1), strDepts)
    
    With rptAdvice
        .Records.DeleteAll
        For i = 1 To rsTmp.RecordCount
            Set objRecord = .Records.Add()
            objRecord.Tag = "0"
            
            Set objItem = objRecord.AddItem("") '选择列
                strTmp = rsTmp!姓名 & ""
                If rsTmp!婴儿姓名 & "" <> "" Then
                    strTmp = strTmp & "之婴(" & rsTmp!婴儿姓名 & ")"
                End If
            Set objItem = objRecord.AddItem(strTmp)
                objItem.Icon = img16.ListImages.Item(IIF(rsTmp!性别 & "" = "男", "Man", "Woman")).Index - 1
            Set objItem = objRecord.AddItem(rsTmp!住院号 & "")
            Set objItem = objRecord.AddItem(rsTmp!床号 & "")
            Set objItem = objRecord.AddItem(rsTmp!性别 & "")
            
                If InStr("," & rsTmp!婴儿年龄 & ",", ",天,") > 0 Then
                    strTmp = rsTmp!年龄 & ""
                Else
                    strTmp = rsTmp!婴儿年龄 & ""
                End If
            Set objItem = objRecord.AddItem(strTmp) '年龄
            
            Set objItem = objRecord.AddItem(rsTmp!科室 & "") '科室
                 strTmp = Format(rsTmp!预约时间, "yyyy-MM-dd HH:mm")
            Set objItem = objRecord.AddItem(strTmp) '预约时间
            Set objItem = objRecord.AddItem(rsTmp!预约设备 & "") '预约设备
            Set objItem = objRecord.AddItem(rsTmp!预约号 & "") '预约号
            Set objItem = objRecord.AddItem(rsTmp!预约项目 & "")  '预约项目
	    Set objItem = objRecord.AddItem(rsTmp!打印时间 & "")
            Set objItem = objRecord.AddItem(rsTmp!打印人 & "")
            
            Set objItem = objRecord.AddItem(rsTmp!病人ID & "")
            Set objItem = objRecord.AddItem(rsTmp!主页ID & "")
            Set objItem = objRecord.AddItem(rsTmp!医嘱ID & "") '医嘱ID
        
            '病人颜色
            objRecord.Item(0).ForeColor = zlDatabase.GetPatiColor(NVL(rsTmp!病人类型))
            For j = 1 To objRecord.Childs.Count - 1
                objRecord.Item(j).ForeColor = objRecord.Item(0).ForeColor
            Next
            objRecord.Item(COL_选择).Icon = img16.ListImages.Item("UnCheck").Index - 1
            objRecord.Tag = "1"
            rsTmp.MoveNext
        Next
        .Populate
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptAdvice.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptAdvice_RowDblClick(rptAdvice.SelectedRows(0), rptAdvice.SelectedRows(0).Record.Item(COL_选择))
        End If
    End If
End Sub

Private Sub rptAdvice_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '如果点击表头的图片，就选中全部
    If Button = 1 Then
        If rptAdvice.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptAdvice.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = COL_选择 Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptAdvice.Columns(COL_选择).Icon = img16.ListImages("UnCheck").Index - 1
                        For i = 0 To rptAdvice.Records.Count - 1
                            rptAdvice.Records(i)(COL_选择).Icon = img16.ListImages("UnCheck").Index - 1
                            rptAdvice.Rows(i).Record.Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptAdvice.Columns(COL_选择).Icon = img16.ListImages("Check").Index - 1
                        For i = 0 To rptAdvice.Records.Count - 1
                            rptAdvice.Records(i)(COL_选择).Icon = -1
                            rptAdvice.Rows(i).Record.Tag = "0"
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptAdvice_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(COL_选择).Icon = -1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(COL_选择).Icon = img16.ListImages.Item("UnCheck").Index - 1
        Row.Record.Tag = "1"
    End If
    rptAdvice.Populate
End Sub

Private Sub SelAllCls(ByVal blnSel As Boolean)
'功能：全选或者全清
'参数：blnSel true -选全，false -全清
    Dim i As Long
     
    If blnSel Then
        rptAdvice.Columns(COL_选择).Caption = "1"
        rptAdvice.Columns(COL_选择).Icon = img16.ListImages("UnCheck").Index - 1
        For i = 0 To rptAdvice.Records.Count - 1
            rptAdvice.Records(i)(COL_选择).Icon = img16.ListImages("UnCheck").Index - 1
            rptAdvice.Rows(i).Record.Tag = "1"
        Next
    Else
        rptAdvice.Columns(COL_选择).Caption = ""
        rptAdvice.Columns(COL_选择).Icon = img16.ListImages("Check").Index - 1
        For i = 0 To rptAdvice.Records.Count - 1
            rptAdvice.Records(i)(COL_选择).Icon = -1
            rptAdvice.Rows(i).Record.Tag = "0"
        Next
    End If
    rptAdvice.Populate
End Sub

Private Function InitUnits() As Boolean
'功能：初始化住院护理病区
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim strUnits As String
    
    On Error GoTo errH
    
    strUnits = GetUser病区IDs
    
    cboUnit.Clear
    
    If InStr(mstrPrivs, "全院病人") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSql = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = strSql & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=A.科室ID)" & _
            " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=A.科室ID)" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSql & ") Group by ID,编码,名称 Order by 编码"
    End If
     
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If InStr(mstrPrivs, "全院病人") > 0 Then
                If rsTmp!ID = UserInfo.部门ID Then '直接所属优先
                    Call cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
                If InStr("," & strUnits & ",", "," & rsTmp!ID & ",") > 0 And cboUnit.ListIndex = -1 Then
                    Call cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            Else '所属缺省病区包含的可能有多个
                If rsTmp!缺省 = 1 And cboUnit.ListIndex = -1 Then
                    Call cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then
        Call cbo.SetIndex(cboUnit.hwnd, 0)
    End If
    
    Call cbo.Locate(cboUnit, mlng病区ID, True)
    
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Private Sub cboUnit_Click()
'功能：刷新界面数据
'说明：从该事件开始会不重复引发相关的数据读取

    If cboUnit.ListIndex = mintPreDept Then Exit Sub
    mintPreDept = cboUnit.ListIndex
    mlng病区ID = Val(cboUnit.ItemData(cboUnit.ListIndex))
    Call LoadDept
    Call LoadAdvice
End Sub

Private Sub ShowAppend(ByVal lng医嘱ID As Long)
'功能：显示指定医嘱的单据附项内容
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, lngIdx As Long
     
    rtfAppend.Text = "": rtfAppend.SelStart = 0
    
    On Error GoTo errH
    
    If lng医嘱ID = 0 Then Exit Sub
    strSql = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order by 排列"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng医嘱ID)
    If Not rsTmp.EOF Then
        With rtfAppend
            Do While Not rsTmp.EOF
                .SelBold = False
                .SelText = IIF(.Text = "", "", vbCrLf) & rsTmp!项目 & "：" & NVL(rsTmp!内容)
                lngIdx = .Find(rsTmp!项目 & "：", , , rtfNoHighlight Or rtfMatchCase)
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(rsTmp!项目 & "：")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
                
                rsTmp.MoveNext
            Loop
            
            rsTmp.MoveFirst
            lngIdx = .Find(rsTmp!项目 & "：", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsTmp!项目 & "：")
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean)
'功能：查找(下一个)病人
'参数：blnNext=是否查找下一个
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    
    If txtFilter.Text = "" Then
        txtFilter.SetFocus
        Exit Sub
    End If
            
    '开始查找行
    If rptAdvice.SelectedRows.Count > 0 Then
        If Not rptAdvice.SelectedRows(0).GroupRow Then
            If Val(rptAdvice.SelectedRows(0).Record(COL_病人ID).value) <> 0 Then blnHave = True
        End If
    End If
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0 'ReportControl的索引从是0开始
    Else
        i = rptAdvice.SelectedRows(0).Index + 1
    End If
    
    '查找病人
    For i = i To rptAdvice.Rows.Count - 1
        With rptAdvice.Rows(i)
            If Not .GroupRow Then
                If mstrFindType = "床号" Then
                    If UCase(Trim(.Record(COL_床号).value)) = UCase(txtFilter.Text) Then Exit For
                ElseIf mstrFindType = "住院号" Then
                    If .Record(COL_住院号).value = txtFilter.Text Then Exit For
                ElseIf mstrFindType = "姓名" Then
                    If .Record(COL_姓名).value Like "*" & txtFilter.Text & "*" Then Exit For
                End If
            End If
        End With
    Next

    If i <= rptAdvice.Rows.Count - 1 Then
        blnReStart = False
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set rptAdvice.FocusedRow = rptAdvice.Rows(i)
        If rptAdvice.Visible Then rptAdvice.SetFocus
    Else
        blnReStart = True
        MsgBox IIF(blnNext, "后面已", "") & "找不到符合条件的病人。", vbInformation, gstrSysName
    End If
End Sub

Private Sub txtDept_KeyPress(KeyAscii As Integer)
    If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtFilter_GotFocus()
    zlControl.TxtSelAll txtFilter
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    Select Case mstrFindType
        Case "住院号"
            If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case "床号"
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case "姓名"
    End Select
    If KeyAscii = 13 Then
        Call ExecuteFindPati
        If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub
