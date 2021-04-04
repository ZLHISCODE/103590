VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockInTendData 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4770
      Index           =   0
      Left            =   195
      ScaleHeight     =   4770
      ScaleWidth      =   9855
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   825
      Width           =   9855
      Begin XtremeReportControl.ReportControl rptData 
         Height          =   2190
         Left            =   435
         TabIndex        =   2
         Top             =   795
         Width           =   3930
         _Version        =   589884
         _ExtentX        =   6932
         _ExtentY        =   3863
         _StockProps     =   0
         ShowGroupBox    =   -1  'True
      End
      Begin VB.Frame fra 
         Height          =   540
         Left            =   15
         TabIndex        =   3
         Top             =   -45
         Width           =   9375
         Begin VB.ComboBox cbo 
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   150
            Width           =   3690
         End
         Begin VB.ComboBox cboBaby 
            Height          =   300
            Left            =   9135
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   165
            Width           =   1350
         End
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "记录范围:"
            Height          =   180
            Left            =   60
            TabIndex        =   6
            Top             =   180
            Width           =   810
         End
      End
      Begin MSComctlLib.ImageList imgData 
         Left            =   3810
         Top             =   4080
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
               Picture         =   "frmDockInTendData.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockInTendData.frx":6862
               Key             =   "体温"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockInTendData.frx":6DFC
               Key             =   "普通"
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgPrint 
      Height          =   1395
      Left            =   10185
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1755
      Visible         =   0   'False
      Width           =   1335
      _cx             =   2355
      _cy             =   2461
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
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   0
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDockInTendData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
''######################################################################################################################
'
Private Enum mCol
    r标志 = 0: rID: r发生时间: r记录项目: r记录数据: r分组: r护士: r登记时间: r病区ID: r记录标记: r病区名:: r签名人: r签名时间: r项目序号: r开始版本: r未记说明: r归档人: r归档时间
    f标志 = 0: fID: f编号: f文件: f日期范围: f病区id: f病区名: f护理级别: f保留
    w标志 = 0: wID: w页面编号: w页面名称: w病历名称: w创建人: w创建时间: w保存人: w完成时间: w当前版本: w签名级别: w当前情况: w归档人: w归档日期: w病区ID: w病区名: w病人状态
End Enum

Private Enum mColWidth
    c标志 = 20: cID = 0: c发生时间 = 110: c记录项目 = 100: c记录数据 = 240: c分组 = 100: c护士 = 60: c登记时间 = 110: c病区id = 0: c记录标记 = 0: c病区名 = 100: c签名人 = 60: c签名时间 = 100: c项目序号 = 0: c开始版本 = 0: c未记说明 = 60: c归档人 = 60: c归档时间 = 110
End Enum

Private mstrPrivs As String                             '当前使用者对本程序(1255)的权限串
Private mlngPatiId As Long                              '病人id
Private mlngPageId As Long                              '主页id
Private mlngDeptId As Long                              '当前操作科室id，如病人科室和当前科室不一致，则不能操作归档外的功能
Private mblnEdit As Boolean                             '是否允许操作，通常由上级程序根据当前操作科室是否当前病人病区决定。
Private mblnDoctorStation As Boolean
Private mintBaby As Integer
Private mblnArchived As Boolean
Private mfrmMain As Object
Private mbytFontSize As Byte
Private WithEvents mfrmCaseTendEdit As frmCaseTendEdit
Attribute mfrmCaseTendEdit.VB_VarHelpID = -1
Private WithEvents mfrmCaseTendEditForBatch As frmCaseTendEditForBatch
Attribute mfrmCaseTendEditForBatch.VB_VarHelpID = -1

Public Event Activate()
Public Event AfterDataChanged()
Public Event AfterArchiveChanged(ByVal blnArchived As Boolean)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

''######################################################################################################################

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-18 15:16
    '问题:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-18 15:16
    '问题:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytSize As Byte
    Dim lngCol As Long
    Dim PATI_COLWIDTH As Variant
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Me.FontSize = mbytFontSize
    Me.FontName = "宋体"
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
            Case UCase("Label")
                objCtrl.FontSize = mbytFontSize
                objCtrl.Height = TextHeight("刘") + 20
            Case UCase("ComboBox")
                objCtrl.FontSize = mbytFontSize
            Case UCase("ReportControl")
                Set CtlFont = objCtrl.PaintManager.CaptionFont
                CtlFont.Size = mbytFontSize
                Set objCtrl.PaintManager.CaptionFont = CtlFont
                
                Set CtlFont = objCtrl.PaintManager.TextFont
                CtlFont.Size = mbytFontSize
                Set objCtrl.PaintManager.TextFont = CtlFont
                PATI_COLWIDTH = Array(c标志, cID, c发生时间, c记录项目, c记录数据, c分组, c护士, c登记时间, c病区id, c记录标记, c病区名, c签名人, r签名时间, r项目序号, r开始版本, r未记说明, r归档人, r归档时间)
                For lngCol = cID To rptData.Columns.Count - 1
                    rptData.Columns.Column(lngCol).Width = BlowUp(CDbl(PATI_COLWIDTH(lngCol)))
                Next lngCol
                '完成列宽的设置
                objCtrl.Redraw
        End Select
    Next
    
    '进行位置调整
    cbo.Top = 150
    cbo.Left = lblData.Left + lblData.Width
    cbo.Width = BlowUp(3690)
    lblData.Top = cbo.Top + (cbo.Height - lblData.Height) \ 2
    cboBaby.Top = cbo.Top
    cboBaby.Width = BlowUp(1350)
    fra.Height = cbo.Top + cbo.Height + 75
    Call picPane_Resize(0)
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '放大：字体，单元格宽度
    BlowUp = dblChange + (dblChange * IIf(mbytFontSize = 12, 1, 0) / 3)
End Function

Private Function ShowOpenedForm() As Boolean
    
    Dim frmTemp As Form
    
    For Each frmTemp In Forms
        
        If frmTemp.Name = "frmCaseTendEdit" Then
            
            ShowSimpleMsg "护理记录编辑窗体已打开，不能重复打开，自动恢复已打开状态！"
            mfrmCaseTendEdit.Show
            
            If mfrmCaseTendEdit.WindowState = 1 Then mfrmCaseTendEdit.WindowState = 0
            mfrmCaseTendEdit.ZOrder 0
            ShowOpenedForm = True
            
            Exit Function
        End If
    Next
End Function

Public Function InitData(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Dim rptCol As ReportColumn
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    
    '------------------------------------------
    '记录数据表设置
    With rptData

        .SetImageList Me.imgData

        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
        End With


        Set rptCol = .Columns.Add(mCol.r标志, "", mColWidth.c标志, False)
        rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter

        Set rptCol = .Columns.Add(mCol.rID, "ID", mColWidth.cID, False): rptCol.Editable = False: rptCol.Groupable = False

        Set rptCol = .Columns.Add(mCol.r发生时间, "发生时间", mColWidth.c发生时间, False): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r记录项目, "记录项目", mColWidth.c记录项目, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r记录数据, "记录数据", mColWidth.c记录数据, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.AutoSize = True
        Set rptCol = .Columns.Add(mCol.r分组, "分组", mColWidth.c分组, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r护士, "护士", mColWidth.c护士, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.r登记时间, "登记时间", mColWidth.c登记时间, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.r病区ID, "科室ID", mColWidth.c病区id, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.r记录标记, "记录标记", mColWidth.c记录标记, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.r病区名, "科室", mColWidth.c病区名, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r签名人, "签名人", mColWidth.c签名人, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r签名时间, "签名时间", mColWidth.c签名时间, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r项目序号, "项目序号", mColWidth.c项目序号, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.r开始版本, "开始版本", mColWidth.c开始版本, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.r未记说明, "未记说明", mColWidth.c未记说明, True):   rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r归档人, "归档人", mColWidth.c归档人, True):   rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.r归档时间, "归档时间", mColWidth.c归档时间, True):   rptCol.Editable = False: rptCol.Groupable = False

    End With

    With cboBaby
        .AddItem "病人本人"
        .ListIndex = 0
    End With

'    If ExecuteCommand("初始控件") = False Or ExecuteCommand("初始数据") = False Then Exit Function
'    Call ExecuteCommand("读注册表")
'    Call ExecuteCommand("控件状态")
    
End Function

Public Function RefreshData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngDeptId As Long, ByVal blnDoctorStation As Boolean, ByVal blnEdit As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：刷新数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    mlngPatiId = lng病人ID
    mlngPageId = lng主页ID
    mblnDoctorStation = blnDoctorStation
    mblnEdit = blnEdit And Not mblnMoved_HL
    mlngDeptId = lngDeptId
    
    cboBaby.Clear
    cboBaby.AddItem "病人本人"
    
    gstrSQL = "Select a.序号,Decode(a.婴儿姓名,Null,NVL(c.姓名,b.姓名) ||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名" & _
        " From 病人信息 b,病案主页 c,病人新生儿记录 a Where b.病人id=c.病人id And a.病人id=c.病人id And a.主页id=c.主页id And c.病人id=[1] And c.主页id=[2]  Order By a.序号"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngPatiId, mlngPageId)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cboBaby.AddItem rs("婴儿姓名").Value
            rs.MoveNext
        Loop
    End If

    cboBaby.ListIndex = 0
    cboBaby.Visible = (cboBaby.ListCount > 1)

    Call zlRefDate(mlngPatiId, mlngPageId)
    
End Function

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strInfo As String
        
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview
        Call zlRptPrint(0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print
        Call zlRptPrint(1)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        Call zlRptPrint(3)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_RowPrint
        Call zlRptPrint(1)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem

        '护理登记：病人ID;主页ID;病区ID;出院;编辑;婴儿
        
        If ShowOpenedForm = False Then
            If mfrmCaseTendEdit Is Nothing Then Set mfrmCaseTendEdit = New frmCaseTendEdit
            If mfrmCaseTendEdit.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & mintBaby & ";2;", 1, mstrPrivs) Then
    '            RaiseEvent AfterDataChanged
    '            cbo.Tag = ""
    '            Call zlRefRec
            End If
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify                    '修改护理记录数据
        
        If ExecuteCommand("修改护理数据") Then
'            cbo.Tag = ""
'            RaiseEvent AfterDataChanged
'            Call zlRefRec
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete

        '删除护理登记
        If rptData.FocusedRow Is Nothing Then Exit Sub
        If rptData.FocusedRow.Record Is Nothing Then Exit Sub

        If MsgBox("确定要删除当前的护理记录吗？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub

        Dim strStart As String
        Dim strEnd As String
        Dim strDate As String

        strDate = rptData.FocusedRow.Record(mCol.r发生时间).Value
        strStart = strDate & ":00"
        strEnd = Format(DateAdd("n", 1, CDate(strDate)), "yyyy-MM-dd HH:mm") & ":00"

        gstrSQL = "ZL_电子护理记录_UPDATE("
        gstrSQL = gstrSQL & mlngPatiId & ","
        gstrSQL = gstrSQL & mlngPageId & ","
        gstrSQL = gstrSQL & mintBaby & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
        gstrSQL = gstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
        gstrSQL = gstrSQL & "1,"
        gstrSQL = gstrSQL & Val(rptData.FocusedRow.Record(mCol.r项目序号).Value) & ","
        gstrSQL = gstrSQL & Val(rptData.FocusedRow.Record(mCol.r记录标记).Value) & ","
        gstrSQL = gstrSQL & "NULL"
        gstrSQL = gstrSQL & ")"

        '执行
        Err = 0: On Error GoTo errHand
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Err = 0: On Error GoTo 0
        RaiseEvent AfterDataChanged
        cbo.Tag = ""
        Call zlRefRec

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search

        '护理记录
        If rptData.FocusedRow Is Nothing Then Exit Sub
        If rptData.FocusedRow.Record Is Nothing Then Exit Sub
        
        If ShowOpenedForm = False Then
            If mfrmCaseTendEdit Is Nothing Then Set mfrmCaseTendEdit = New frmCaseTendEdit
            Call mfrmCaseTendEdit.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & mintBaby & ";2;" & CStr(rptData.FocusedRow.Record.Item(mCol.r发生时间).Value), 5)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Sign


        '护理记录
        If rptData.FocusedRow Is Nothing Then Exit Sub
        If rptData.FocusedRow.Record Is Nothing Then Exit Sub
        
        If ShowOpenedForm = False Then
            
            If mfrmCaseTendEdit Is Nothing Then Set mfrmCaseTendEdit = New frmCaseTendEdit
            If mfrmCaseTendEdit.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & mintBaby & ";2;" & CStr(rptData.FocusedRow.Record.Item(mCol.r发生时间).Value), 3) Then
    '            cbo.Tag = ""
    '            RaiseEvent AfterDataChanged
    '            Call zlRefRec
            End If
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_SignEarse

        '护理记录
        If rptData.FocusedRow Is Nothing Then Exit Sub
        If rptData.FocusedRow.Record Is Nothing Then Exit Sub
        
        If ShowOpenedForm = False Then
            If mfrmCaseTendEdit Is Nothing Then Set mfrmCaseTendEdit = New frmCaseTendEdit
            If mfrmCaseTendEdit.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & mintBaby & ";2;" & CStr(rptData.FocusedRow.Record.Item(mCol.r发生时间).Value), 4) Then
    '            cbo.Tag = ""
    '            RaiseEvent AfterDataChanged
    '            Call zlRefRec
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10

        If MsgBox("需要将该病人本次住院所有护理记录归档吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then

            Dim strNow As String

            strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            gstrSQL = "Zl_电子护理记录_Archive(" & mlngPatiId & "," & mlngPageId & "," & mintBaby & ",'" & gstrUserName & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            mblnArchived = True
                        
            cbo.Tag = ""
            Call zlRefRec
            
            Err = 0: On Error GoTo 0

        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_UnArchive

        If mblnArchived Then
            If MsgBox("需要撤销该病人本次住院所有已归档护理记录吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then

                gstrSQL = "Zl_电子护理记录_UnArchive(" & mlngPatiId & "," & mlngPageId & "," & mintBaby & ")"
                Err = 0: On Error GoTo errHand
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                
                mblnArchived = False
                cbo.Tag = ""
                Call zlRefRec
                
                Err = 0: On Error GoTo 0

            End If
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MarkMap
        '体温作图：病人ID;主页ID;病区ID;出院;编辑;婴儿
        If Not CreateBodyEditor Then Exit Sub
        If gobjBodyEditor.GetTendBody.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";0;1;" & mintBaby, 2, mstrPrivs) Then

            Call zlRefDate(mlngPatiId, mlngPageId)

            RaiseEvent AfterDataChanged
            cbo.Tag = ""
            Call zlRefRec
        End If

    Case conMenu_File_PrintDayDetail        '批量录入
        If mfrmCaseTendEditForBatch Is Nothing Then Set mfrmCaseTendEditForBatch = New frmCaseTendEditForBatch
        Call mfrmCaseTendEditForBatch.ShowMe(Me, mlngDeptId, mstrPrivs)
        
    End Select
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
LL:
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
'    Dim lngCount As Long, blnFinished As Boolean, lngMaxVersion As Long, eSignLevel As EPRSignLevelEnum

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML
        Control.Enabled = (rptData.Records.Count > 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        Control.Enabled = (rptData.Records.Count > 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem
        
        Control.Visible = (InStr(1, mstrPrivs, "护理记录登记") > 0 And mblnDoctorStation = False)
        Control.Enabled = (Control.Visible And mblnEdit And mlngPatiId > 0 And mblnArchived = False And Not mblnMoved_HL)
                
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
    
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And mblnArchived = False And Not mblnMoved_HL)

        Control.Visible = (InStr(1, mstrPrivs, "护理记录登记") > 0 And mblnDoctorStation = False)
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "护理记录登记") > 0)
        If Control.Enabled Then Control.Enabled = Not (Me.rptData.FocusedRow Is Nothing)
        If Control.Enabled Then Control.Enabled = Not Me.rptData.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value >= 0)
        If (InStr(1, mstrPrivs, "他人护理记录") = 0) Then
            If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.r护士).Value = gstrUserName)
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And mblnArchived = False And Not mblnMoved_HL)

        Control.Visible = (InStr(1, mstrPrivs, "护理记录登记") > 0 And mblnDoctorStation = False)

        If Control.Enabled Then Control.Enabled = Control.Visible
        If Control.Enabled Then Control.Enabled = Not (Me.rptData.FocusedRow Is Nothing)
        If Control.Enabled Then Control.Enabled = Not Me.rptData.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value >= 0)
        If (InStr(1, mstrPrivs, "他人护理记录") = 0) Then
            If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.r护士).Value = gstrUserName)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search

        Control.Visible = (mblnDoctorStation = False)
        Control.Enabled = (mlngPatiId > 0 And Control.Visible)
        If Control.Enabled Then Control.Enabled = Not (Me.rptData.FocusedRow Is Nothing)
        If Control.Enabled Then Control.Enabled = Not Me.rptData.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value >= 0)
        If Control.Enabled Then Control.Enabled = (rptData.FocusedRow.Record(mCol.r开始版本).Value > 0)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Sign

        Control.Visible = (InStr(1, mstrPrivs, "护理记录签名") > 0 And mblnDoctorStation = False)
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible And mblnArchived = False And Not mblnMoved_HL)
        If Control.Enabled Then Control.Enabled = Not (Me.rptData.FocusedRow Is Nothing)
        If Control.Enabled Then Control.Enabled = Not Me.rptData.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value >= 0)
        If Control.Enabled Then Control.Enabled = (rptData.FocusedRow.Record(mCol.r签名人).Value = "")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_SignEarse
        Control.Visible = (InStr(1, mstrPrivs, "取消记录签名") > 0 And mblnDoctorStation = False)
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible And mblnArchived = False And Not mblnMoved_HL)
        If Control.Enabled Then Control.Enabled = Not (Me.rptData.FocusedRow Is Nothing)
        If Control.Enabled Then Control.Enabled = Not Me.rptData.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value >= 0)
        If Control.Enabled Then Control.Enabled = (rptData.FocusedRow.Record(mCol.r开始版本).Value > 0)
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10

        Control.Visible = (InStr(1, mstrPrivs, "护理记录归档") > 0 And mblnDoctorStation = False And mblnArchived = False)
        Control.Enabled = Control.Visible And mblnEdit And Not mblnMoved_HL

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_UnArchive

        Control.Visible = (InStr(1, mstrPrivs, "取消记录归档") > 0 And mblnDoctorStation = False And mblnArchived)
        Control.Enabled = Control.Visible And mblnEdit And Not mblnMoved_HL

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup

        Control.Visible = (mblnDoctorStation = False And (InStr(1, mstrPrivs, "体温单作图") > 0 Or InStr(1, mstrPrivs, "护理记录登记") > 0))
        Control.Enabled = Control.Visible
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MarkMap

        Control.Visible = (InStr(1, mstrPrivs, "体温单作图") > 0 And mblnDoctorStation = False)
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible And mblnArchived = False And Not mblnMoved_HL)
    
    Case conMenu_File_PrintDayDetail
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And mblnArchived = False And Not mblnMoved_HL)

        Control.Visible = (InStr(1, mstrPrivs, "护理记录登记") > 0 And mblnDoctorStation = False)
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "护理记录登记") > 0) And Not mblnDoctorStation
    End Select
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        
               
            
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"

        
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"

        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新状态"
        

        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
                
        
    '------------------------------------------------------------------------------------------------------------------
    Case "修改护理数据"
    
        '护理登记
        If rptData.FocusedRow Is Nothing Then Exit Function
        If rptData.FocusedRow.Record Is Nothing Then Exit Function
        
        If ShowOpenedForm = False Then
            If mfrmCaseTendEdit Is Nothing Then Set mfrmCaseTendEdit = New frmCaseTendEdit
            ExecuteCommand = mfrmCaseTendEdit.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & mintBaby & ";2;" & CStr(rptData.FocusedRow.Record.Item(mCol.r发生时间).Value) & ";" & CStr(rptData.FocusedRow.Record.Item(mCol.rID).Value), 2, mstrPrivs)
        End If
        
        Exit Function
        
    End Select

    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:

End Function

Private Sub cbo_Click()

    Call zlRefRec

End Sub

Private Sub cboBaby_Click()

    If mintBaby = cboBaby.ListIndex Then Exit Sub
    mintBaby = cboBaby.ListIndex

    Call zlRefDate(mlngPatiId, mlngPageId)

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
    If Not mfrmCaseTendEdit Is Nothing Then Unload mfrmCaseTendEdit
    If Not mfrmCaseTendEditForBatch Is Nothing Then Unload mfrmCaseTendEditForBatch
    Set mfrmCaseTendEdit = Nothing
    Set mfrmCaseTendEditForBatch = Nothing
    
End Sub

Private Sub mfrmCaseTendEdit_AfterDataChanged()
    RaiseEvent AfterDataChanged
    cbo.Tag = ""
    Call zlRefDate(mlngPatiId, mlngPageId)
End Sub

Private Sub mfrmCaseTendEditForBatch_AfterDataChanged()
    RaiseEvent AfterDataChanged
    cbo.Tag = ""
    
    Call zlRefDate(mlngPatiId, mlngPageId)
End Sub

Private Sub rptData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not (rptData.FocusedRow Is Nothing) Then
            Call rptData_RowDblClick(rptData.FocusedRow, rptData.FocusedRow.Record.Item(mCol.r发生时间))
        End If
    End If
End Sub

Private Sub rptData_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

Private Sub rptData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

    If Not (rptData.FocusedRow Is Nothing) Then
        
        RaiseEvent RowDblClick(Row, Item)

    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next

    RaiseEvent Activate
End Sub

Private Function zlRefDate(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim cbrItem As CommandBarControl
    Dim intCount As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim strEnterDate As String
    Dim intCol As Integer
    Dim strCaption As String
    Dim strParameter As String
    Dim strSvrCaption As String
    Dim strNow As String
    Dim strCut As String
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lnglast科室id As Long
    Dim intSvrDate As Integer
    Dim blnData As Boolean '是否存在老板数据
    
    strCut = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    If cbo.ListIndex >= 0 Then intSvrDate = cbo.ItemData(cbo.ListIndex)
    
    cbo.Clear
    cbo.Tag = ""
    cbo.AddItem "所有记录"
    cbo.ItemData(cbo.NewIndex) = 0

    '------------------------------------------------------------------------------------------------------------------
                
    strSQL = "Select 入院时间, 出院时间, 1 + Nvl(Round((b.出院时间 - b.入院时间) / 7),-1) As 页数" & vbNewLine & _
                "  from (Select Min(发生时间) as 入院时间," & vbNewLine & _
                "               Max(发生时间) as 出院时间" & vbNewLine & _
                "          From 病人护理记录" & vbNewLine & _
                "         Where 病人ID = [1] And 主页ID = [2]) b"
    If mblnMoved_HL Then
        strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
        strSQL = Replace(strSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng病人ID, lng主页ID)
    If rsTmp.BOF Then Exit Function
    
    If NVL(rsTmp!入院时间) <> "" Then blnData = True
    
    '
    '------------------------------------------------------------------------------------------------------------------
                
'    strSQL = "Select 1 As 开始页码,1 + Round((a.终止时间 - a.开始时间) / 7) As 结束页码," & vbNewLine & _
'                "       科室id,c.名称," & vbNewLine & _
'                "       开始时间," & vbNewLine & _
'                "       终止时间" & vbNewLine & _
'                "  from (Select 科室id," & vbNewLine & _
'                "               Min(发生时间) as 开始时间," & vbNewLine & _
'                "               Max(发生时间) as 终止时间" & vbNewLine & _
'                "          From 病人护理记录" & vbNewLine & _
'                "         Where 病人ID = [1] And 主页ID = [2]" & vbNewLine & _
'                "         Group by 科室id) a," & vbNewLine & _
'                "       (Select Min(开始时间) as 入院时间" & vbNewLine & _
'                "          From 病人变动记录" & vbNewLine & _
'                "         Where 开始时间 is Not Null And 病人ID = [1] And 主页ID = [2]) b,部门表 c Where c.ID=a.科室id " & vbNewLine & _
'                " order by a.开始时间"

    strSQL = "Select 1 As 开始页码, 1 + Round((a.终止时间 - a.开始时间) / 7) As 结束页码, 开始时间, 终止时间" & vbNewLine & _
             "   From (Select Min(发生时间) As 开始时间, Max(发生时间) As 终止时间" & vbNewLine & _
             "          From 病人护理记录" & vbNewLine & _
             "          Where 病人id = [1] And 主页id = [2]) A," & vbNewLine & _
             "        (Select Min(开始时间) As 入院时间" & vbNewLine & _
             "          From 病人变动记录" & vbNewLine & _
             "          Where 开始时间 Is Not Null And 病人id = [1] And 主页id = [2]) B" & vbNewLine & _
             "   Order By a.开始时间"
    If mblnMoved_HL Then
        strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
        strSQL = Replace(strSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng病人ID, lng主页ID)

    strEnterDate = Format(rsTmp!入院时间, "yyyy-MM-dd HH:mm:ss")
    For lngLoop = 0 To rsTmp("页数").Value - 1

        strDateFrom = Format(rsTmp("入院时间").Value + 7 * lngLoop, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("入院时间").Value + 7 * (lngLoop + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
        If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
        End If

        If strDateFrom < Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then

            If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss")

            rs.Filter = ""
            rs.Filter = "开始页码<=" & lngLoop + 1 & " And 结束页码>=" & lngLoop + 1
            rs.Sort = "开始时间"
            If rs.RecordCount > 0 Then rs.MoveFirst
            For intCol = 1 To rs.RecordCount

                If strDateFrom < Format(rs("开始时间").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strTmp = Format(rs("开始时间").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strTmp = strDateFrom
                End If

                If strDateTo > Format(rs("终止时间").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strCaption = Format(rs("终止时间").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strCaption = strDateTo
                End If

                strCaption = Format(strTmp, "yyyy年MM月dd日") & " ～ " & Format(strCaption, "yyyy年MM月dd日")

                cbo.AddItem strCaption
                cbo.ItemData(cbo.NewIndex) = intCol

                rs.MoveNext

            Next
        End If

    Next
    
    If intSvrDate > 0 Then
        Call zlControl.CboLocate(cbo, intSvrDate)
        If cbo.ListIndex = -1 Then cbo.ListIndex = cbo.ListCount - 1
    Else
        cbo.ListIndex = cbo.ListCount - 1
    End If
    
    If mblnEdit = True Then
        '41778,刘鹏飞,2012-09-06
        '如果病人老板和新版数据都已经存在，不做任何限制。如果只有新板数据，没有老版。则老板不能添加文件。
        '婴儿应该和母亲使用同一套系统。
        strSQL = "Select 1 From 病人护理文件 A Where a.病人id = [1] And a.主页id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        If rsTmp.RecordCount > 0 And blnData = False Then
            mblnEdit = False
        End If
    End If
    
    zlRefDate = True
End Function

Private Function zlRefRec() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
Dim rsTemp As New ADODB.Recordset
Dim strTmp As String
Dim strStart As String
Dim strEnd As String
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem

    On Error GoTo errHand

    If cbo.Tag = cbo.Text Then
        zlRefRec = True
        Exit Function
    End If

    cbo.Tag = cbo.Text

    mblnArchived = False

    gstrSQL = "Select 归档人,归档时间 From 病人护理记录 Where 病人id=[1] And 主页id=[2] And Nvl(婴儿,0)=[3] And RowNum<2 And 归档人 Is Not Null"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mintBaby)
    If rsTemp.BOF = False Then
        mblnArchived = True
    End If
    
    RaiseEvent AfterArchiveChanged(mblnArchived)
    '

    '------------------------------------------------------------------------------------------------------------------
    '护理数据刷新
    If cbo.ItemData(cbo.ListIndex) = 0 Then
        gstrSQL = "Select Decode(f.记录人,Null,0,1) As 开始版本,c.Id, e.记录人 As 签名人,e.项目名称 As 签名时间,Nvl(c.未记说明,c.记录内容) As 内容,c.记录标记,r.发生时间, c.项目名称, Decode(c.项目序号,1,Decode(c.记录标记,1,Null,Decode(c.体温部位,Null,'腋温',c.体温部位)||':'),Null)||c.记录内容 || c.项目单位 || Decode(c.记录标记, 1, Decode(c.项目序号,1,'(物理降温)',Null), Null) As 记录内容," & _
                "        c.项目分组, c.项目序号, c.记录人, nvl(c.修改时间,r.保存时间) AS 保存时间, r.科室id As 病区id, d.名称 As 病区名,c.未记说明,r.归档人,r.归档时间 " & _
                " From 病人护理记录 r, 病人护理内容 c, 部门表 d,病人护理内容 e,病人护理内容 f " & _
                " Where r.Id = c.记录id And r.科室id = d.Id And r.病人id = [1] And r.主页id = [2] And c.记录类型 = 1 And  Nvl(r.婴儿,0)=[3] And c.终止版本 Is Null And r.ID=e.记录id(+) And e.记录类型(+)=5 And Nvl(r.最后版本,1)=Nvl(e.开始版本(+),1)  And f.记录id(+)=c.记录id And f.记录类型(+)=5 And Nvl(f.开始版本(+),1)=1 " & _
                " Order By r.发生时间 Desc"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
            gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mintBaby)
    Else

        strTmp = Trim(Mid(cbo.Text, InStr(cbo.Text, ")") + 1))
        strStart = Format(Trim(Mid(strTmp, 1, InStr(strTmp, "～") - 1)), "yyyy-MM-dd")
        strEnd = Format(Trim(Mid(strTmp, InStr(strTmp, "～") + 1)), "yyyy-MM-dd") & " 23:59:59"

        gstrSQL = "Select Decode(f.记录人,Null,0,1) As 开始版本,c.Id, e.记录人 As 签名人,e.项目名称 As 签名时间,c.记录内容 As 内容,c.记录标记,r.发生时间, c.项目名称, Decode(c.项目序号,1,Decode(c.记录标记,1,Null,Decode(c.体温部位,Null,'腋温',c.体温部位)||':'),Null)||c.记录内容 || c.项目单位 || Decode(c.记录标记, 1, Decode(c.项目序号,1,'(物理降温)',Null), Null) As 记录内容," & _
                "        c.项目分组, c.项目序号, c.记录人, nvl(c.修改时间,r.保存时间) AS 保存时间, r.科室id As 病区id, d.名称 As 病区名,c.未记说明,r.归档人,r.归档时间 " & _
                " From 病人护理记录 r, 病人护理内容 c, 部门表 d,病人护理内容 e,病人护理内容 f " & _
                " Where r.Id = c.记录id And r.科室id = d.Id And r.病人id = [1] And r.主页id = [2] And c.记录类型 = 1 And  Nvl(r.婴儿,0)=[5]  And 发生时间 Between [3] And [4] And c.终止版本 Is Null And r.ID=e.记录id(+) And e.记录类型(+)=5 And Nvl(r.最后版本,1)=Nvl(e.开始版本(+),1) And f.记录id(+)=c.记录id And f.记录类型(+)=5 And Nvl(f.开始版本(+),1)=1 " & _
                " Order By r.发生时间 Desc"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
            gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, CDate(strStart), CDate(strEnd), mintBaby)
    End If

    rptData.Records.DeleteAll
    With rsTemp
        Do While Not rsTemp.EOF
            Set rptRcd = rptData.Records.Add()
            Set rptItem = rptRcd.AddItem(""): rptItem.Icon = 0
            rptRcd.AddItem CStr("" & !ID)
            rptRcd.AddItem Format(!发生时间, "yyyy-MM-dd hh:mm")
            rptRcd.AddItem CStr("" & !项目名称)

            strTmp = CStr("" & !记录内容)
            Select Case rsTemp("项目序号").Value
            Case 9
                If Right(rsTemp("内容").Value, 1) = "C" Then
                    strTmp = CStr("" & !内容)
                End If
            Case 10
                If zlCommFun.NVL(rsTemp("内容").Value) <> "" Then
                    If Right(rsTemp("内容").Value, 2) = "/E" Then
                        strTmp = CStr("" & !内容)
                    ElseIf Right(rsTemp("内容").Value, 1) = "E" Then
                        strTmp = CStr("" & !内容)
                    ElseIf Right(rsTemp("内容").Value, 1) = "*" Then
                        strTmp = CStr("" & !内容)
                    End If
                End If
            End Select
            
            If zlCommFun.NVL(rsTemp("未记说明").Value) <> "" Then
                rptRcd.AddItem CStr(rsTemp("未记说明").Value)
            Else
                        
                If zlCommFun.NVL(rsTemp("内容").Value) = "" Then
                    rptRcd.AddItem ""
                Else
                    rptRcd.AddItem strTmp
                End If
            End If

            rptRcd.AddItem CStr("" & !项目分组)
            rptRcd.AddItem CStr("" & !记录人)
            rptRcd.AddItem Format(!保存时间, "yyyy-MM-dd hh:mm")
            rptRcd.AddItem CStr("" & !病区ID)
            rptRcd.AddItem CStr("" & !记录标记)
            rptRcd.AddItem CStr("" & !病区名)
            rptRcd.AddItem CStr("" & !签名人)
            rptRcd.AddItem Format(!签名时间, "yyyy-MM-dd hh:mm")
            rptRcd.AddItem CStr("" & !项目序号)
            rptRcd.AddItem Val(zlCommFun.NVL(!开始版本))
            rptRcd.AddItem CStr("" & !未记说明)
            rptRcd.AddItem CStr("" & !归档人)
            If IsNull(!归档时间) Then
                rptRcd.AddItem ""
            Else
                rptRcd.AddItem Format(!归档时间, "yyyy-MM-dd hh:mm")
            End If
                        
            .MoveNext
        Loop
    End With
    Me.rptData.Populate
    If Me.rptData.Records.Count > 0 Then Set Me.rptData.FocusedRow = rptData.Rows(0)

    zlRefRec = True

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '       strSubhead，打印的副标题
    '-------------------------------------------------
Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
Dim rsTemp As New ADODB.Recordset
    
    '获得基本信息
    Dim strSubhead As String
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select b.住院号, NVL(b.姓名,a.姓名) 姓名 From 病人信息 a,病案主页 b Where a.病人id=b.病人id And b.病人id = [1] And b.主页id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
    If Not rsTemp.EOF Then
        strSubhead = "住院号:" & rsTemp!住院号 & "  姓名:" & rsTemp!姓名
    Else
        strSubhead = ""
    End If
    Err = 0: On Error GoTo 0

    If Me.rptData.Records.Count = 0 Then Exit Sub
    If zlReportToVSFlexGrid(Me.vfgPrint, Me.rptData) = False Then Exit Sub

    Call vfgPrint.AutoSize(0, vfgPrint.Cols - 1)

    Set objPrint.Body = Me.vfgPrint
    objPrint.Title.Text = "护理记录数据清单"


    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add(strSubhead)
    Call objAppRow.Add("第" & mlngPageId & "次住院")
    Call objPrint.UnderAppRows.Add(objAppRow)

    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)

    Me.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.Tag = ""
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
    
        fra.Move 0, -90, picPane(Index).Width
        
        cboBaby.Move fra.Width - cboBaby.Width, cboBaby.Top
        
        rptData.Move 15, fra.Top + fra.Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (fra.Top + fra.Height + 15) - 15
    End Select
End Sub
