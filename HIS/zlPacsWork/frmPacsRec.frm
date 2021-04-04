VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPACSRec 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6270
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   315
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraFee 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   5760
      Width           =   7935
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   " 申请报告"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   45
         Width           =   810
      End
   End
   Begin VB.Frame fraSplit1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      Top             =   5640
      Width           =   7110
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Height          =   1260
      Left            =   0
      TabIndex        =   5
      Top             =   6120
      Width           =   7260
      _cx             =   12806
      _cy             =   2222
      Appearance      =   1
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPacsRec.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
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
      Begin MSComctlLib.ImageList imgFlag 
         Left            =   765
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   8
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPacsRec.frx":009B
               Key             =   "未填"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPacsRec.frx":05B5
               Key             =   "已填"
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkHistory 
      Caption         =   "显示历史病历"
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwFile 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      SmallIcons      =   "imgFlag"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "日期"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "病历记录"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "医生"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "状态"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "类别"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsRec.frx":0ACF
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPACSRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COL_F申请 = 0 '标志列
Private Const COL_F报告 = 1
Private Const COL_NO = 2 '可见列
Private Const COL_医嘱内容 = 3
Private Const COL_单据 = 4
Private Const COL_申请人 = 5
Private Const COL_申请时间 = 6
Private Const COL_发送时间 = 7
Private Const COL_报告人 = 8
Private Const COL_报告时间 = 9
Private Const COL_医嘱ID = 10 '隐藏列
Private Const COL_诊疗项目ID = 11
Private Const COL_单据ID = 12
Private Const COL_编号 = 13
Private Const COL_申请项 = 14
Private Const COL_申请ID = 15
Private Const COL_报告项 = 16
Private Const COL_报告ID = 17
Private Const COL_记录性质 = 18
Private Const COL_前提ID = 19
Private Const COL_执行状态 = 20

Private PatientID As Long '病人ID
Private PageID As Variant    '主页ID
Private AreaID As Long, DeptID As Long, OffHosp As Boolean
Private mlngAdviceID As Long, blnShowAll As Boolean
Private aFrmEdit() As Form '编辑窗口数组
Private mblnMoved As Boolean

Public mstrPrivs As String
Public WithEvents mfrmParent As Form
Attribute mfrmParent.VB_VarHelpID = -1
Private Const COLOR_归档 As Long = &H8000&
Private Const COLOR_作废 As Long = &HFF&
Private Const COLOR_正常 As Long = &H80000008
Private Const COLOR_其他 As Long = &H808080
'刷新窗体显示内容
Public Sub zlRefresh(ByVal lngPatientID As Long, ByVal varPageID As Variant, _
    Optional ByVal lngAdviceID As Long = 0, Optional ByVal ifShowAll As Boolean = True, _
    Optional ifShowHistory As Boolean = False, Optional ByVal blnMoved As Boolean = False)
'lngPatientID:病人ID（0=没指定病人）
'strCheckID:挂号单
    PatientID = lngPatientID: PageID = varPageID
    mlngAdviceID = lngAdviceID: blnShowAll = ifShowAll
    chkHistory.Value = IIf(ifShowHistory, 1, 0)
    mblnMoved = blnMoved
    
    ShowFile chkHistory
    LoadReport
End Sub
'执行菜单命令
Public Sub zlMenuClick(mnuClick As Menu)
    Dim strMenu As String
    
    '增加病历
    If UCase(mnuClick.Name) = "FILELIST" Then
        AddFile mnuClick
        Exit Sub
    End If
    If UCase(mnuClick.Name) = "REQLIST" Then
        FuncAddRequest CLng(mnuClick.Tag)
        Exit Sub
    End If
    
    If mnuClick.Caption Like "*(&*)*" Then
        strMenu = Split(mnuClick.Caption, "(&")(0)
    Else
        strMenu = mnuClick.Caption
    End If
    Select Case strMenu
        Case "修改病历"
            EditFile
        Case "删除病历"
            DeleteFile
        Case "修改申请单"
            FuncWriteRequest
        Case "删除申请单"
            FuncDeleteRequest
        Case "打印通知单"
            FuncPrintRequest
        Case "查阅报告"
            FuncViewReport
        Case "预览报告"
            FuncPrintReport 1
        Case "打印报告"
            FuncPrintReport 2
        Case "影像对比"
            ViewImage
    End Select
End Sub
Public Sub zlButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "病历"
            Me.PopupMenu mfrmParent.mnuPFileFunc(0)
        Case "删病历"
            DeleteFile
        Case "病历修改"
            EditFile
    End Select
End Sub
Public Sub zlItemRef()

End Sub
Public Sub zlExcel()

End Sub
Public Sub zlPreview()
    If PatientID > 0 And Not lvwFile.SelectedItem Is Nothing Then _
        PreviewPatientFile mfrmParent, IIf(TypeName(PageID) = "String", 1, 2), False, 1, PatientID, PageID, False, 0, 1
End Sub
Public Sub zlPrint()
    Dim btOption As Byte
    Dim bPrtCurrFile As Boolean, bPrtPatiInfo As Boolean, lngBeginY As Long, iBeginPage As Integer
    Dim rsTmp As New ADODB.Recordset, lngFileSeq As Long
    Dim strSQL As String
    If PatientID = 0 Or lvwFile.SelectedItem Is Nothing Then Exit Sub
    
    
    bPrtCurrFile = True: bPrtPatiInfo = False
    lngBeginY = 0: iBeginPage = 1
    btOption = PrintOptionSetup_Patient(mfrmParent, True, bPrtCurrFile, bPrtPatiInfo, lngBeginY, iBeginPage, CLng(Mid(lvwFile.SelectedItem.Key, 4)))
    If btOption = 0 Then Exit Sub
    
    strSQL = "Select Seq From (Select RowNum As Seq,ID From (Select ID From 病人病历记录 a,病历文件目录 b where a.文件ID = b.ID " + _
        " And a.病人ID=" & PatientID & IIf(TypeName(PageID) = "String", _
        " And a.挂号单='" & PageID & "' And a.病历种类=1", _
        " And a.主页ID=" & PageID & " And a.病历种类=2") & _
        " Order By b.新页 desc,a.书写日期)) Where ID= [1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(lvwFile.SelectedItem.Key, 4))

    If rsTmp.EOF Then
        lngFileSeq = 1
    Else
        lngFileSeq = rsTmp(0)
    End If
    Select Case btOption
        Case 1
            PreviewPatientFile mfrmParent, IIf(TypeName(PageID) = "String", 1, 2), bPrtCurrFile, lngFileSeq, PatientID, PageID, bPrtPatiInfo, lngBeginY * 56.7, iBeginPage
        Case 2
            PrintPatientFile mfrmParent, IIf(TypeName(PageID) = "String", 1, 2), bPrtCurrFile, lngFileSeq, PatientID, PageID, bPrtPatiInfo, lngBeginY * 56.7, iBeginPage
    End Select
End Sub
Public Sub zlPrintSetup()
    PrintSetup_Patient mfrmParent
End Sub
Private Sub chkHistory_Click()
    On Error Resume Next
    ShowFile chkHistory: Me.lvwFile.SetFocus
End Sub

'新增病历文件
Private Sub AddFile(mnufile As Menu)
    Dim iNum As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If PatientID = 0 Then Exit Sub
    On Error Resume Next
    iNum = -1: iNum = UBound(aFrmEdit)
    On Error GoTo EditFileError
    '判断前提和是否重复书写
    strSQL = "Select * From 病历文件目录 Where ID= [1] "
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(mnufile.Tag, 2))
    
    If rsTmp.EOF Then Exit Sub
    
    If Not IsNull(rsTmp("书写")) Then
        If rsTmp("书写") = 0 Then
            '不能重复书写的病历，检查是否已书写该病历
            strSQL = "Select Count(*) From 病人病历记录 Where 病人ID= [1] " & _
                IIf(TypeName(PageID) = "String", _
                " And 挂号单= [2] ", _
                " And 主页ID= [2] ") & _
                " And 作废日期 Is Null And 文件ID = [3] "
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, PatientID, PageID, Mid(mnufile.Tag, 2))

            If rsTmp(0) > 0 Then MsgBox "该病历已经存在，不能重复书写", vbExclamation, gstrSysName: Exit Sub
        Else
        End If
    End If
    
    ReDim Preserve aFrmEdit(iNum + 1)
    EditPatientFile "", CStr(PatientID), CStr(PageID), IIf(TypeName(PageID) = "String", 0, 1), Mid(mnufile.Tag, 2), , Me, aFrmEdit(UBound(aFrmEdit)), , IIf(TypeName(PageID) = "String", 1, 2), , mlngAdviceID
    Exit Sub
EditFileError:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    mfrmParent.SetFocus: zlCommFun.PressKey CByte(KeyCode)
End Sub

Private Sub Form_Load()
    lvwFile.ListItems.Add , , "Temp", , 1
    lvwFile.ListItems.Clear
    
    InitBillTable
    
    mfrmParent.mnuFileExcel.Visible = False
    PatientID = -1: PageID = -1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.fraSplit1.Top > Me.ScaleHeight Then Me.fraSplit1.Top = Me.ScaleHeight - 1590
    
    With lvwFile
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
        .Height = Me.fraSplit1.Top - .Top
    End With
    With Me.fraSplit1
        .Left = 0: .Width = Me.ScaleWidth - .Left
    End With
    With Me.fraFee
        .Left = 0: .Top = fraSplit1.Top + fraSplit1.Height
        .Width = Me.ScaleWidth - .Left
    End With
    With Me.vsBill
        .Left = 0: .Top = fraFee.Top + fraFee.Height
        .Width = Me.ScaleWidth - .Left: .Height = Me.ScaleHeight - .Top
    End With
End Sub

'显示病人病历文件记录
Private Sub ShowFile(ByVal ShowHistory As Boolean)
    Dim rsTmp As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim i As Integer, lngColor As Long
    Dim strSQL As String
    
    lvwFile.ListItems.Clear
    '如果没有病人则不再处理
    If PatientID = 0 Then
        ShowMenu
        Exit Sub
    End If
    If ShowHistory Then
        strSQL = "Select ID,书写日期,nvl(病历名称,'未命名') As 病历名称,nvl(书写人,'未命名') As 书写人,nvl(挂号单,'0') As 挂号单," + _
            "Decode(作废日期,Null,Decode(归档日期,Null,' ','归档'),'作废') As 状态,Decode(病历种类,-2,'麻醉记录','入院病历') As 类别,医嘱ID From 病人病历记录 " + _
            "Where 病人ID= [1]  And " & IIf(TypeName(PageID) = "String", "病历种类=1 ", "病历种类 In (2,-2) ") & _
            IIf(mlngAdviceID = 0 Or blnShowAll, "", "And 医嘱ID= [3]  ") & _
            "And 作废日期 Is Null Order By 书写日期 Desc"
    Else
        strSQL = "Select ID,书写日期,nvl(病历名称,'未命名') As 病历名称,nvl(书写人,'未命名') As 书写人,nvl(挂号单,'0') As 挂号单," + _
            "Decode(作废日期,Null,Decode(归档日期,Null,' ','归档'),'作废') As 状态,Decode(病历种类,-2,'麻醉记录','入院病历') As 类别,医嘱ID From 病人病历记录 " + _
            "Where 病人ID= [1] And " & IIf(TypeName(PageID) = "String", "挂号单= [2] And 病历种类=1 ", "主页ID= [2] And 病历种类 In (2,-2) ") & _
            IIf(mlngAdviceID = 0 Or blnShowAll, "", "And 医嘱ID= [3] ") & _
            "And 作废日期 Is Null Order By 书写日期 Desc"
    End If
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
    End If
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, PatientID, PageID, mlngAdviceID)
    
    Do While Not rsTmp.EOF
        Set tmpItem = lvwFile.ListItems.Add(, "Key" & rsTmp("ID"), rsTmp("病历名称"))
        
        tmpItem.Tag = Nvl(rsTmp("医嘱ID"), "0")
        tmpItem.SubItems(1) = IIf(IsNull(rsTmp("书写日期")), "院外", rsTmp("书写日期"))
        tmpItem.SubItems(2) = rsTmp("病历名称")
        tmpItem.SubItems(3) = rsTmp("书写人")
        tmpItem.SubItems(4) = rsTmp("状态")
        tmpItem.SubItems(5) = rsTmp("类别")
        With tmpItem.ListSubItems
            Select Case rsTmp("状态")
                Case "归档"
                    lngColor = COLOR_归档
                Case "作废"
                    lngColor = COLOR_作废
                Case Else
                    lngColor = COLOR_正常
            End Select
            If lngColor = COLOR_正常 And CLng(tmpItem.Tag) <> mlngAdviceID Then lngColor = COLOR_其他
            For i = 1 To lvwFile.ColumnHeaders.Count - 1
                .Item(i).ForeColor = lngColor
            Next
        End With
        
        rsTmp.MoveNext
    Loop
    ShowMenu
End Sub

Private Sub ShowMenu()
    Dim blnEnabled As Boolean
    On Error Resume Next
    
    blnEnabled = Not (Me.lvwFile.SelectedItem Is Nothing)
    mfrmParent.mnuPFileFunc(0).Enabled = Not (TypeName(PageID) = "String" Or PatientID = 0)
    mfrmParent.mnuPFileFunc(1).Enabled = blnEnabled
    mfrmParent.mnuPFileFunc(2).Enabled = blnEnabled
    mfrmParent.tbrMain.Buttons("病历").Enabled = mfrmParent.mnuPFileFunc(0).Enabled
    mfrmParent.tbrMain.Buttons("病历修改").Enabled = mfrmParent.mnuPFileFunc(1).Enabled
    mfrmParent.tbrMain.Buttons("删病历").Enabled = mfrmParent.mnuPFileFunc(2).Enabled
    
    blnEnabled = Not (Val(vsBill.TextMatrix(vsBill.Row, COL_医嘱ID)) = 0)
    mfrmParent.mnuReqFunc(0).Enabled = Not (TypeName(PageID) = "String")
    mfrmParent.mnuReqFunc(1).Enabled = blnEnabled
    mfrmParent.mnuReqFunc(2).Enabled = blnEnabled
    mfrmParent.mnuReqFunc(4).Enabled = blnEnabled
    mfrmParent.mnuReqFunc(6).Enabled = blnEnabled
    mfrmParent.mnuReqFunc(7).Enabled = blnEnabled
    mfrmParent.mnuReqFunc(8).Enabled = blnEnabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, iNum As Integer
    
    '卸载病历编辑窗口
    On Error Resume Next
    iNum = -1: iNum = UBound(aFrmEdit)
    For i = 0 To iNum
        Unload aFrmEdit(0)
    Next
    
    Call SaveWinState(Me, App.ProductName, mfrmParent.Name)
End Sub

Private Sub fraSplit1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    fraSplit1.BackColor = RGB(0, 0, 0)
    On Error Resume Next
    If fraSplit1.Top + y < 2000 Then
        fraSplit1.Top = 2000
    ElseIf Me.ScaleHeight - fraSplit1.Top - y < 2000 Then
        fraSplit1.Top = Me.ScaleHeight - 2000
    Else
        fraSplit1.Top = fraSplit1.Top + y
    End If
End Sub

Private Sub fraSplit1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    fraSplit1.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub lvwFile_DblClick()
    EditFile
End Sub

Private Sub lvwFile_KeyDown(KeyCode As Integer, Shift As Integer)
    mfrmParent.SetFocus: zlCommFun.PressKey CByte(KeyCode)
End Sub

Private Sub lvwFile_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And mfrmParent.mnuPFile.Visible And mfrmParent.mnuPFile.Enabled Then Me.PopupMenu mfrmParent.mnuPFile
End Sub
'删除病历文件
Private Sub DeleteFile()
    Dim iCurrIndex As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lngEditHWnd As Long
    Dim strSQL As String
    
    If lvwFile.SelectedItem Is Nothing Then Exit Sub
    
    strSQL = "Select * From 病人病历记录 Where ID= [1] And Not 归档日期 Is Null"
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(lvwFile.SelectedItem.Key, 4))
    
    If Not rsTmp.EOF Then
        MsgBox "该病历文件已归档，不能删除！", vbInformation, gstrSysName
        Exit Sub
    End If
    If CLng(lvwFile.SelectedItem.Tag) <> mlngAdviceID Then
        MsgBox "只能删除本次检查书写的病历文件！", vbInformation, gstrSysName
        Exit Sub
    End If
    If mblnMoved Then
        MsgBox "当前病人的病历已转入备份，不能执行本操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    lngEditHWnd = GetEditWindow(CLng(Mid(lvwFile.SelectedItem.Key, 4)))
    
    If lngEditHWnd > 0 Then
        MsgBox "该病历正在编辑，不能删除", vbExclamation, gstrSysName
        Call ShowWindow(lngEditHWnd, SW_RESTORE)
        Call BringWindowToTop(lngEditHWnd)
    Else
        If MsgBox("是否删除该病历文件？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            With lvwFile
                strSQL = "ZL_病人病历_DELETE(" + Mid(.SelectedItem.Key, 4) + ")"
                ExecuteProc strSQL, Me.Caption
'                zlDatabase.ExecuteProcedure "ZL_病人病历_DELETE(" + Mid(.SelectedItem.Key, 4) + ")", ""
                
                iCurrIndex = .SelectedItem.Index
                .ListItems.Remove iCurrIndex
            End With
            
            ShowMenu
        End If
    End If
End Sub
'修改病历文件
Private Sub EditFile()
    Dim iNum As Integer, lngEditHWnd As Long
    
    If lvwFile.SelectedItem Is Nothing Then Exit Sub
    If Not mfrmParent.mnuPFile.Visible Then Exit Sub
    
    If mblnMoved Then
        MsgBox "当前病人的病历已转入备份，不能执行本操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lngEditHWnd = GetEditWindow(CLng(Mid(lvwFile.SelectedItem.Key, 4)))
    If lngEditHWnd = 0 Then
        On Error Resume Next
        iNum = -1: iNum = UBound(aFrmEdit)
        On Error GoTo EditFileError
        ReDim Preserve aFrmEdit(iNum + 1)
        EditPatientFile Mid(lvwFile.SelectedItem.Key, 4), CStr(PatientID), CStr(PageID), IIf(TypeName(PageID) = "String", 0, 1), , , Me, aFrmEdit(UBound(aFrmEdit)), _
             (Len(Trim(lvwFile.SelectedItem.SubItems(4))) = 0) And _
             mfrmParent.mnuPFile.Visible And _
             CLng(lvwFile.SelectedItem.Tag) = mlngAdviceID, IIf(TypeName(PageID) = "String", 1, 2), , mlngAdviceID
    Else
        Call ShowWindow(lngEditHWnd, SW_RESTORE)
        Call BringWindowToTop(lngEditHWnd)
    End If
    Exit Sub
EditFileError:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub
'病历文件归档
Private Sub CreateFolder()
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim lngColor As Long
    Dim strSQL As String
    
    If lvwFile.SelectedItem Is Nothing Then Exit Sub
    strSQL = "Select * From 病人病历记录 Where ID= [1] And Not 归档日期 Is Null"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(lvwFile.SelectedItem.Key, 4))
    
    If Not rsTmp.EOF Then
        MsgBox "该病历文件已归档！", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("病历归档后将不能删除和修改，继续吗？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then
        With lvwFile
            strSQL = "ZL_病人病历_归档(" + Mid(.SelectedItem.Key, 4) + ",'" + UserInfo.姓名 + "')"
            ExecuteProc strSQL, Me.Caption
'            zlDatabase.ExecuteProcedure "ZL_病人病历_归档(" + Mid(.SelectedItem.Key, 4) + ",'" + UserInfo.姓名 + "')", ""
        End With
        
        lvwFile.SelectedItem.SubItems(4) = "归档"
        With lvwFile.SelectedItem.ListSubItems
            lngColor = COLOR_归档
            For i = 1 To lvwFile.ColumnHeaders.Count - 1
                .Item(i).ForeColor = lngColor
            Next
        End With
        
        ShowMenu
    End If
End Sub
'病历文件作废
Private Sub UndoFile()
    Dim iCurrIndex As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If lvwFile.SelectedItem Is Nothing Then Exit Sub
    
    strSQL = "Select * From 病人病历记录 Where ID=  [1] And Not 归档日期 Is Null And 作废日期 Is Null"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(lvwFile.SelectedItem.Key, 4))

    If rsTmp.EOF Then
        MsgBox "该病历文件不能作废！", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("确认将该份病历作废吗？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then
        With lvwFile
            strSQL = "ZL_病人病历_作废(" + Mid(.SelectedItem.Key, 4) + ",'" + UserInfo.姓名 + "')"
            ExecuteProc strSQL, Me.Caption
'            zlDatabase.ExecuteProcedure "ZL_病人病历_作废(" + Mid(.SelectedItem.Key, 4) + ",'" + UserInfo.姓名 + "')", ""
            
            iCurrIndex = .SelectedItem.Index
            .ListItems.Remove iCurrIndex
        End With
        
        ShowMenu
    End If
End Sub

'根据病历记录ID获取其编辑窗口的hwnd,为0则表示该病历当前未编辑
Private Function GetEditWindow(ByVal lngFileID As Long) As Long
    Dim i As Integer, iNum As Integer
    
    On Error Resume Next
    iNum = -1: iNum = UBound(aFrmEdit)
    For i = 0 To iNum
        If aFrmEdit(i).Tag = lngFileID Then GetEditWindow = aFrmEdit(i).Hwnd: Exit For
    Next
End Function
'病历编辑窗口关闭后的处理。由病历编辑调用。
Public Sub EditFile_UnLoad(ByVal lngHwnd As Long)
    Dim i As Integer, iNum As Integer, iTmpIndex As Integer
    
    '从编辑窗口数组中删除当前关闭的窗口
    On Error Resume Next
    iNum = -1: iNum = UBound(aFrmEdit)
    For i = 0 To iNum
        If aFrmEdit(i).Hwnd = lngHwnd Then Exit For
    Next
    iTmpIndex = i
    For i = iTmpIndex + 1 To iNum
        Set aFrmEdit(i - 1) = aFrmEdit(i)
    Next
    Set aFrmEdit(iNum) = Nothing
    If iNum = 0 Then
        Erase aFrmEdit
    Else
        ReDim Preserve aFrmEdit(iNum - 1)
    End If
    
    ShowFile chkHistory
End Sub

Private Sub mfrmParent_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub FuncPrintRequest()
'功能：打印通知单
    Dim strBill As String
    
    If PatientID = 0 Then Exit Sub
    
    With vsBill
        If Val(.TextMatrix(.Row, COL_医嘱ID)) = 0 Then Exit Sub
        
        '如果无申请内容则不必
        If Val(.TextMatrix(.Row, COL_申请项)) = 0 Then
            MsgBox "该单据不需要填写申请，不能打印通知单。", vbInformation, gstrSysName
            Exit Sub
        End If
        '如果未填写申请则不允许
        If Val(.TextMatrix(.Row, COL_申请ID)) = 0 Then
            MsgBox "该单据还没有填写申请，不能打印通知单。", vbInformation, gstrSysName
            Exit Sub
        End If
        '如果已填写报告则不允许
        If Val(.TextMatrix(.Row, COL_报告ID)) <> 0 Then
            MsgBox "该单据已经填写报告，不能打印通知单。", vbInformation, gstrSysName
            Exit Sub
        End If
        If .TextMatrix(.Row, COL_前提ID) <> mlngAdviceID Then
            MsgBox "只能打印本次检查填写的申请！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strBill = "ZLCISBILL" & Format(.TextMatrix(.Row, COL_编号), "00000") & "-1"
        If ReportPrintSet(gcnOracle, glngSys, strBill, mfrmParent) Then
            Call ReportOpen(gcnOracle, glngSys, strBill, mfrmParent, "NO=" & .TextMatrix(.Row, COL_NO), "性质=" & Val(.TextMatrix(.Row, COL_记录性质)), 2)
        End If
    End With
End Sub

Private Sub FuncPrintReport(ByVal PrtMode As Integer)
'功能：打印报告单
    Dim strBill As String
    
    If PatientID = 0 Then Exit Sub
    
    With vsBill
        If Val(.TextMatrix(.Row, COL_医嘱ID)) = 0 Then Exit Sub
        '如果无报告内容则不必
        If Val(.TextMatrix(.Row, COL_报告项)) = 0 Then
            MsgBox "该单据不需要填写报告，不能打印报告单。", vbInformation, gstrSysName
            Exit Sub
        End If

        '如果未填写报告则不允许
        If Val(.TextMatrix(.Row, COL_报告ID)) = 0 Then
            MsgBox "该单据还没有填写报告，不能打印报告单。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(.TextMatrix(.Row, COL_执行状态)) <> 1 Then
            MsgBox "该报告尚未审核，不能打印。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        PrintDiagRpt_New .TextMatrix(.Row, COL_报告ID), mfrmParent, PrtMode, picBuffer, mblnMoved
    End With
End Sub

Private Sub FuncAddRequest(ByVal lng单据ID As Long)
'功能：新增申请单
    If PatientID = 0 Then Exit Sub
    If lng单据ID = 0 Then Exit Sub
    
    '调用接口
    Call AddRequest(mfrmParent, PatientID, PageID, lng单据ID, False, , , mlngAdviceID)
    If True Then
        Call LoadReport
    End If
End Sub

Private Sub FuncWriteRequest()
'功能：填写辅诊申请
    If PatientID = 0 Then Exit Sub
    With vsBill
        If Val(.TextMatrix(.Row, COL_医嘱ID)) = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, COL_申请项)) = 0 Then
            MsgBox "该单据不需要填写申请。", vbInformation, gstrSysName
            Exit Sub
        End If
        If .TextMatrix(.Row, COL_NO) <> "" Then
            MsgBox "该医嘱已经发送，不能再填写申请。", vbInformation, gstrSysName
            Exit Sub
        End If
'        If .TextMatrix(.Row, COL_前提ID) <> mlngAdviceID Then
'            MsgBox "只能修改本次检查填写的申请！", vbInformation, gstrSysName
'            Exit Sub
'        End If
        
        '填写申请:医嘱ID,单据ID,申请ID,医嘱内容
        Call EditRequest(Me, Val(.TextMatrix(.Row, COL_医嘱ID)), Val(.TextMatrix(.Row, COL_单据ID)), Val(.TextMatrix(.Row, COL_申请ID)), .TextMatrix(.Row, COL_医嘱内容), .TextMatrix(.Row, COL_前提ID) <> mlngAdviceID, DataMoved:=mblnMoved)
        If True Then
            Call LoadReport
        End If
    End With
End Sub

Private Sub FuncViewReport()
'功能：填写辅诊申请
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If PatientID = 0 Then Exit Sub
    With vsBill
        If Val(.TextMatrix(.Row, COL_报告ID)) = 0 Then Exit Sub
        '判断是否紧急打印过的报告,此类报告可以查看,
        strSQL = "SELECT a.是否打印, b.紧急标志 FROM 影像检查记录 a ,病人医嘱记录 b where a.医嘱id=b.id and a.医嘱id = [1]"
        Set rsTmp = OpenSQLRecord(strSQL, "查阅报告", CLng(vsBill.TextMatrix(vsBill.Row, COL_医嘱ID)))
        If rsTmp.EOF Then Exit Sub
        If (Val(.TextMatrix(.Row, COL_执行状态)) <> 1) And (Nvl(rsTmp(0), 0) = 0 Or Nvl(rsTmp(1), 0) = 0) Then
            MsgBox "该报告尚未审核，不能查阅。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '填写申请:医嘱ID,单据ID,申请ID,医嘱内容
        Call EditReport(Me, .TextMatrix(.Row, COL_NO), Val(.TextMatrix(.Row, COL_记录性质)), _
            Val(.TextMatrix(.Row, COL_单据ID)), Val(.TextMatrix(.Row, COL_报告ID)), .TextMatrix(.Row, COL_医嘱内容), True, DataMoved:=mblnMoved)
        If True Then
            Call LoadReport
        End If
    End With
End Sub

Private Sub FuncDeleteRequest()
'功能：删除当前申请单
    Dim strSQL As String, lngRow As Long
        
    If PatientID = 0 Then Exit Sub
    With vsBill
        If Val(.TextMatrix(.Row, COL_单据ID)) = 0 Then Exit Sub
        '具有申请附项的单据
        If Val(.TextMatrix(.Row, COL_申请项)) = 0 Then
            MsgBox "单据[" & .TextMatrix(.Row, COL_单据) & "]没有需要申请的内容。", vbInformation, gstrSysName
            Exit Sub
        End If
        '已填写申请单
        If Val(.TextMatrix(.Row, COL_申请ID)) = 0 Then
            MsgBox "单据[" & .TextMatrix(.Row, COL_单据) & "]没有填写申请部份的内容。", vbInformation, gstrSysName
            Exit Sub
        End If
        '已发送后不能删除(可以通过医嘱作废)
        If .TextMatrix(.Row, COL_NO) <> "" Then
            MsgBox "该医嘱已经发送，对应的申请单不能再删除。", vbInformation, gstrSysName
            Exit Sub
        End If
        If .TextMatrix(.Row, COL_前提ID) <> mlngAdviceID Then
            MsgBox "只能删除本次检查填写的申请！", vbInformation, gstrSysName
            Exit Sub
        End If
        If mblnMoved Then
            MsgBox "当前病人的申请已转入备份，不能执行本操作！", vbInformation, gstrSysName
            Exit Sub
        End If
        If MsgBox("确实要删除申请单[" & .TextMatrix(.Row, COL_单据) & "]吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '已校对对的在过程中检查
        strSQL = "zl_病人医嘱记录_Delete(" & Val(.TextMatrix(.Row, COL_医嘱ID)) & ",1)"
    End With
    
    '删除申请单
    On Error GoTo errH
    gcnOracle.BeginTrans
    Call ExecuteProc(strSQL, Me.Caption)
    gcnOracle.CommitTrans
    On Error GoTo 0
        
    '更新界面
    With vsBill
        lngRow = .Row
        .RemoveItem .Row
        If .Rows = .FixedRows Then
            .Rows = .FixedRows + 1
        End If
        If lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        Call .ShowCell(.Row, .Col)
    End With
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitBillTable()
'功能：初始化单据清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "单据号,810,1;医嘱内容,3000,1;单据,1800,1;申请人,850,1;" & _
        "申请时间,1080,1;发送时间,1080,1;报告人,850,1;报告时间,1080,1;" & _
        "医嘱ID;诊疗项目ID;单据ID;编号;申请项;申请ID;报告项;报告ID;记录性质;前提ID;执行状态"
    arrHead = Split(strHead, ";")
    With vsBill
        .Clear
        .FixedRows = 1: .FixedCols = 2
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .ColWidth(0) = 11 * Screen.TwipsPerPixelX
        .ColWidth(1) = 11 * Screen.TwipsPerPixelX
    End With
End Sub

Private Function LoadReport() As Boolean
'功能：根据当前病人医嘱读取可以填写的申请单或报告单
    Dim rsBill As New ADODB.Recordset
    Dim strSQL As String, strBill As String
    Dim strKey As String, lngPreRow As Long, i As Long
    Dim strSqlWhere As String
    
    If PatientID = 0 Then
        With vsBill
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
        End With
        Exit Function
    End If
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    With vsBill
        If Val(.TextMatrix(.Row, COL_医嘱ID)) <> 0 Then
            strKey = Val(.TextMatrix(.Row, COL_医嘱ID)) & "_" & .TextMatrix(.Row, COL_NO)
        End If
        .Redraw = flexRDNone
        .Rows = .FixedRows
    End With
    
    '诊疗单据具有申请或报告附项的医嘱
    strBill = "Select A.ID as 医嘱ID," & _
        " B.病历文件ID as 单据ID,D.编号,D.名称,D.说明," & _
        " Max(Decode(C.填写时机,1,1,0)) as 申请项," & _
        " Max(Decode(C.填写时机,2,1,0)) as 报告项" & _
        " From 病人医嘱记录 A,诊疗单据应用 B,病历文件组成 C,病历文件目录 D" & _
        " Where A.诊疗项目ID=B.诊疗项目ID And B.应用场合=" & IIf(TypeName(PageID) = "String", 1, 2) & _
        " And B.病历文件ID=C.病历文件ID And B.病历文件ID=D.ID" & _
        " And A.病人ID= [1] " & IIf(TypeName(PageID) = "String", " And 挂号单= [2] ", " And A.主页ID= [2] ") & _
        " And (A.诊疗类别 Not IN('5','6','7') And A.相关ID is NULL" & _
        "   Or A.诊疗类别='C' And A.相关ID is Not NULL)" & _
        IIf(mlngAdviceID = 0 Or blnShowAll, "", " And A.前提ID= [3] ") & _
        " Group by A.ID,B.病历文件ID,D.编号,D.名称,D.说明"
    
    '与药品相关的医嘱
    strSqlWhere = "Select Distinct 相关ID From 病人医嘱记录" & _
        " Where 病人ID= [1] " & IIf(TypeName(PageID) = "String", " And 挂号单= [2] ", " And A.主页ID= [2] ") & _
        " And 诊疗类别 IN('5','6','7')"
        
    '医嘱对应的单据清单(包括待安排医嘱,不含不发送的叮嘱),至少包含一种单据附项
    '未发送的医嘱显示一条,已发送的一次发送显示一条(包括只有申请项的)
    strSQL = _
        " Select A.ID,A.诊疗项目ID,A.医嘱内容,B.发送时间,B.NO,B.记录性质," & _
        " A.申请ID,B.报告ID,C.编号,C.名称,C.单据ID,C.申请项,C.报告项," & _
        " X.书写人 as 申请人,X.书写日期 as 申请时间," & _
        " Y.书写人 as 报告人,Y.书写日期 as 报告时间,Nvl(A.前提ID,0) As 前提ID,Nvl(B.执行状态,0) As 执行状态,A.诊疗类别,B.发送号" & _
        " From 病人医嘱记录 A,病人医嘱发送 B,(" & strBill & ") C,病人病历记录 X,病人病历记录 Y" & _
        " Where A.病人ID= [1] " & IIf(TypeName(PageID) = "String", " And A.挂号单= [2] ", " And A.主页ID= [2] ") & _
        " And (A.诊疗类别 Not IN('5','6','7') And A.相关ID is NULL" & _
        "   Or A.诊疗类别='C' And A.相关ID is Not NULL)" & _
        " And A.ID Not IN(" & strSqlWhere & ") And A.医嘱状态<>4 And Nvl(A.执行性质,0)<>0" & _
        " And A.ID=B.医嘱ID(+) And A.ID=C.医嘱ID And (C.申请项=1 Or C.报告项=1)" & _
        " And A.申请ID=X.ID(+) And B.报告ID=Y.ID(+)" & _
        " Order by Nvl(B.发送时间,A.开嘱时间) Desc,A.序号"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
    End If
    
    '医嘱内容,NO,单据,申请人,申请时间,发送时间,报告人,报告时间
    '医嘱ID;诊疗项目ID;单据ID;编号;申请项;申请ID;报告项;报告ID;记录性质
    Set rsBill = OpenSQLRecord(strSQL, Me.Caption, PatientID, PageID, mlngAdviceID)
   
    With vsBill
        .Rows = .FixedRows + rsBill.RecordCount
        For i = 1 To rsBill.RecordCount
            .TextMatrix(i, COL_医嘱内容) = rsBill!医嘱内容
            .TextMatrix(i, COL_NO) = Nvl(rsBill!NO)
            .TextMatrix(i, COL_单据) = rsBill!名称
            .TextMatrix(i, COL_申请人) = Nvl(rsBill!申请人)
            .TextMatrix(i, COL_申请时间) = Format(Nvl(rsBill!申请时间), "MM-dd HH:mm")
            .TextMatrix(i, COL_发送时间) = Format(Nvl(rsBill!发送时间), "MM-dd HH:mm")
            .TextMatrix(i, COL_报告人) = Nvl(rsBill!报告人)
            .TextMatrix(i, COL_报告时间) = Format(Nvl(rsBill!报告时间), "MM-dd HH:mm")
            .TextMatrix(i, COL_医嘱ID) = rsBill!ID
            .TextMatrix(i, COL_诊疗项目ID) = rsBill!诊疗项目ID
            .TextMatrix(i, COL_单据ID) = rsBill!单据ID
            .TextMatrix(i, COL_编号) = rsBill!编号
            .TextMatrix(i, COL_申请项) = Nvl(rsBill!申请项, 0)
            .TextMatrix(i, COL_申请ID) = Nvl(rsBill!申请ID, 0)
            .TextMatrix(i, COL_报告项) = Nvl(rsBill!报告项, 0)
            .TextMatrix(i, COL_报告ID) = Nvl(rsBill!报告ID, 0)
            .TextMatrix(i, COL_记录性质) = Nvl(rsBill!记录性质)
            .TextMatrix(i, COL_前提ID) = rsBill!前提ID
            .TextMatrix(i, COL_执行状态) = rsBill!执行状态
            
            .Cell(flexcpData, i, COL_诊疗项目ID) = Nvl(rsBill!诊疗类别)
            .Cell(flexcpData, i, COL_发送时间) = rsBill!发送号
            
            '申请与报告的标识
            If rsBill!申请项 = 1 Then
                If Not IsNull(rsBill!申请ID) Then
                    Set .Cell(flexcpPicture, i, COL_F申请) = imgFlag.ListImages("已填").Picture
                Else
                    Set .Cell(flexcpPicture, i, COL_F申请) = imgFlag.ListImages("未填").Picture
                End If
            End If
            If rsBill!报告项 = 1 Then
                If Not IsNull(rsBill!报告ID) Then
                    Set .Cell(flexcpPicture, i, COL_F报告) = imgFlag.ListImages("已填").Picture
                Else
                    Set .Cell(flexcpPicture, i, COL_F报告) = imgFlag.ListImages("未填").Picture
                End If
            End If
            
            '定位到先前行
            If Val(.TextMatrix(i, COL_医嘱ID)) & "_" & .TextMatrix(i, COL_NO) = strKey Then
                lngPreRow = i
            End If
            
            If .TextMatrix(i, COL_前提ID) <> mlngAdviceID Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = COLOR_其他
            End If
            
            rsBill.MoveNext
        Next
        If .Rows = .FixedRows Then
            .Rows = .FixedRows + 1
        Else
            .AutoSize COL_医嘱内容
        End If
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
        
        .Col = COL_NO
        .Row = IIf(lngPreRow <> 0, lngPreRow, .FixedRows)
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
    End With
    ShowMenu
    
    Screen.MousePointer = 0
    LoadReport = True
    Exit Function
errH:
    Screen.MousePointer = 0
    vsBill.Redraw = flexRDDirect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And mfrmParent.mnuReq.Visible And mfrmParent.mnuReq.Enabled Then Me.PopupMenu mfrmParent.mnuReq
End Sub

Private Sub vsBill_RowColChange()
    '忽略观片菜单
    On Error Resume Next
    mfrmParent.mnuReqFunc(10).Enabled = (vsBill.Cell(flexcpData, vsBill.Row, COL_诊疗项目ID) = "D" And _
        Val(vsBill.TextMatrix(vsBill.Row, COL_执行状态)) = 1)
End Sub

Private Sub ViewImage()
'功能：调用观片站
    Dim aFiles() As String
    Dim objPacsCore As Object
    Dim strFTPHost As String, strDicomPath As String, strLocalPath As String
    Dim strFTPUser As String, strFtpPwd As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strCachePath As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strCheckUID As String
    
    On Error GoTo DBError
    strSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "观片处理", CLng(vsBill.TextMatrix(vsBill.Row, COL_医嘱ID)))
    If rsTmp.EOF Then Exit Sub
    
    strCachePath = App.Path & "\TmpImage\"
    If Not objFileSystem.FolderExists(strCachePath) Then objFileSystem.CreateFolder strCachePath
    
    strCheckUID = Nvl(rsTmp(0))
    aFiles = GetAllImageFiles(strCheckUID, , mblnMoved, strFTPHost, strDicomPath, _
        strLocalPath, strFTPUser, strFtpPwd)
    If UBound(aFiles) > 0 Then
        Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
        objPacsCore.CallOpenViewerCache aFiles, mfrmParent, strCachePath & strLocalPath, strFTPHost & strDicomPath, mstrPrivs, strCheckUID, strFTPHost, strDicomPath, gcnOracle, strFTPUser, strFtpPwd, True
        Set objPacsCore = Nothing
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
