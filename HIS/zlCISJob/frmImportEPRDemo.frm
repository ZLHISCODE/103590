VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportEPRDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "范文导入..."
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmImportEPRDemo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3195
      Left            =   120
      TabIndex        =   3
      Top             =   390
      Width           =   6705
      _Version        =   589884
      _ExtentX        =   11827
      _ExtentY        =   5636
      _StockProps     =   0
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.TextBox txtFilt 
      Height          =   300
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   1
      Top             =   3690
      Width           =   3330
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5685
      TabIndex        =   5
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "打开(&O)"
      Height          =   350
      Left            =   5685
      TabIndex        =   4
      Top             =   3675
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3375
      Top             =   4350
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
            Picture         =   "frmImportEPRDemo.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportEPRDemo.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportEPRDemo.frx":0EBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblList 
      AutoSize        =   -1  'True
      Caption         =   "当前“###”可用范文(&L):"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2070
   End
   Begin VB.Shape shpList 
      BorderColor     =   &H00FFC0C0&
      Height          =   3225
      Left            =   105
      Top             =   375
      Width           =   6735
   End
   Begin VB.Label lblFilt 
      AutoSize        =   -1  'True
      Caption         =   "简码过滤(&F)"
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   3750
      Width           =   990
   End
End
Attribute VB_Name = "frmImportEPRDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngFileID As Long          '定义文件id
Private mlngPatient As Long         '病人id，在病人病历编辑时，用来确定条件示范是否满足
Private mlngVisit As Long           '主页id或挂号单ID
Private mstrLike As String
'Private mintPower As Integer    '权限
Private mblnOK As Boolean        '返回
Private mrsInfo As Recordset     '返回范文内容记录集
Private mstrDbOwner As String                '当前数据库所有者（不同模块可能不一样）


'################################################################################################################
'## 功能：  装入指定ID的范文列表
'##
'## 参数：  frmParent       :父窗体
'##
'## 返回：  返回选中的范文ID
'################################################################################################################
Public Function ShowMe(frmParent As Object, ByVal lngFileID As Long, ByVal lngPatient As Long, ByVal lngVisit As Long, ByRef rsInfo As Recordset) As Long
    Dim strCaption As String
    Dim lngCount As Long

    strCaption = "快捷病历"
    mlngFileID = lngFileID
    mlngPatient = lngPatient
    mlngVisit = lngVisit
    mstrDbOwner = GetDbOwner(glngSys)
    lngCount = FillEPRDemos
    On Error Resume Next
    If lngCount = 1 Then
        cmdOK_Click
        If mblnOK Then ShowMe = Me.rptList.FocusedRow.Record.Item(1).Value: Set rsInfo = mrsInfo
        Unload Me
        Exit Function
    ElseIf lngCount = 0 Then
        MsgBox "没有可用的范文。", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    On Error GoTo 0
    Me.lblList.Caption = "当前“" & strCaption & "”可用范文(&L):"
    Me.Show vbModal, frmParent
    If mblnOK Then ShowMe = Me.rptList.FocusedRow.Record.Item(1).Value: Set rsInfo = mrsInfo
    Unload Me
End Function

Private Function FillEPRDemos() As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim strSQL As String
    
    strSQL = "Select l.Id, l.编号, l.名称, l.简码, Nvl(l.分类,'未分类') As 分类,l.说明, l.通用级" & vbNewLine & _
            "From 病历范文目录 l, Table(Cast(f_Segment_Usable([1], [2], [3], 0) As " & mstrDbOwner & ".t_Dic_Rowset)) u" & vbNewLine & _
            "Where l.文件id = [1] And Nvl(l.性质, 0) = [4] And l.Id = To_Number(u.编码)"
'    Select Case mintPower
'    Case 0
'    Case 1
'        gstrSQL = gstrSQL & " And" & vbNewLine & _
'                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
'                "      L.通用级 In (1, 2) And" & vbNewLine & _
'                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"
'    Case Else
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 = 1 And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      L.通用级 = 2 And L.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
'    End Select
    
    err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngFileID, mlngPatient, mlngVisit, 0)
    
    Me.rptList.Records.DeleteAll
    If rsTemp.EOF Then cmdOK.Enabled = False
    Do While Not rsTemp.EOF
        Set rptRcd = Me.rptList.Records.Add()
        Set rptItem = rptRcd.AddItem(CStr(Nvl(rsTemp!通用级, 0))): rptItem.Icon = rptItem.Value
        rptRcd.AddItem CStr(rsTemp!ID)
        rptRcd.AddItem CStr(rsTemp!分类)
        rptRcd.AddItem CStr(rsTemp!编号)
        rptRcd.AddItem CStr(rsTemp!名称)
        rptRcd.AddItem CStr("" & rsTemp!简码)
        rptRcd.AddItem CStr("" & rsTemp!说明)
        rsTemp.MoveNext
    Loop
    Me.rptList.Populate
    If Me.rptList.Rows.Count > 1 And Me.rptList.FocusedRow Is Nothing Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(1)
    End If
    
    FillEPRDemos = Me.rptList.Records.Count
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    FillEPRDemos = Me.rptList.Records.Count
End Function

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim blnNoSelected As Boolean
        
    If Me.rptList.FocusedRow.GroupRow Then
        blnNoSelected = True
    ElseIf Me.rptList.FocusedRow Is Nothing Then
        blnNoSelected = True
    End If
    
    If blnNoSelected Then
        MsgBox "没有选中范文，不能打开！", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnOK = True: Me.Hide
    '读取范文内容
    Set mrsInfo = LoadDemoInfo(Val(Me.rptList.FocusedRow.Record.Item(1).Value))
End Sub

Private Function LoadDemoInfo(ByVal lngFileID As Long) As Recordset
'功能：读取指定的范文内容
    Dim strSQL As String
    
    strSQL = "Select A.预制提纲id, B.内容文本, B.ID" & vbNewLine & _
            "From 病历范文内容 A, 病历范文内容 B" & vbNewLine & _
            "Where A.文件id = [1] And A.对象类型 = 1 And A.预制提纲id+0 In(-10,5,2,6,3)" & vbNewLine & _
            "      And B.父id = A.ID And B.对象类型 = 2 Order By A.预制提纲id, B.对象序号"
    Set LoadDemoInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngFileID)

End Function

Private Sub Form_Load()
    Dim rptCol As ReportColumn

    mstrLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    With Me.rptList
        Set rptCol = .Columns.Add(0, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(1, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(2, "分类", 50, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(3, "编号", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(4, "名称", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(5, "简码", 60, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(6, "说明", 200, True): rptCol.Editable = False: rptCol.Groupable = False
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
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(2)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(3)
        
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.Record Is Nothing Then Exit Sub
    
    cmdOK_Click
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdOK_Click
End Sub

Private Sub txtFilt_GotFocus()
    Me.txtFilt.SelStart = 0: Me.txtFilt.SelLength = 4000
    Call ZLCommFun.OpenIme(False)
End Sub

Private Sub txtFilt_KeyPress(KeyAscii As Integer)
    Dim rptRcd As ReportRecord
    
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        For Each rptRcd In Me.rptList.Records
            If Trim(Me.txtFilt.Text) = "" Then
                rptRcd.Visible = True
            Else
                rptRcd.Visible = (rptRcd(5).Value Like IIf(mstrLike <> "", "*", "") & Trim(Me.txtFilt.Text) & "*")
            End If
        Next
        Me.rptList.Populate
        If Me.rptList.Rows.Count > 0 And Me.rptList.FocusedRow Is Nothing Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
        Call txtFilt_GotFocus
        Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        If KeyAscii = Asc("*") Or KeyAscii = Asc("?") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Public Function GetDbOwner(ByVal lngSys As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    GetDbOwner = ""
    err = 0: On Error GoTo errHand
    strSQL = "Select 所有者 From Zlsystems Where 编号 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetDbOwner", lngSys)
    If rsTemp.RecordCount <> 0 Then GetDbOwner = "" & rsTemp!所有者
    rsTemp.Close
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
