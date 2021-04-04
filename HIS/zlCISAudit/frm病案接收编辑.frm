VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm病案接收编辑 
   Caption         =   "病案接收编辑"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11430
   Icon            =   "frm病案接收编辑.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   11430
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraCmd 
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   5400
      Width           =   11415
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   3120
         TabIndex        =   18
         Top             =   285
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   10200
         TabIndex        =   17
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "确定(&O)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   9120
         TabIndex        =   16
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&F)"
         Height          =   350
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打印(&P)"
         Height          =   350
         Left            =   8040
         TabIndex        =   13
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   2520
         TabIndex        =   19
         Top             =   325
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H80000004&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5355
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton cmd运送人 
         Height          =   300
         Left            =   11010
         Picture         =   "frm病案接收编辑.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chkInput 
         Caption         =   "输入住院号"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox txtOuter 
         Height          =   300
         Left            =   6960
         TabIndex        =   5
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtInput 
         Height          =   300
         Left            =   1440
         MaxLength       =   18
         TabIndex        =   4
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox txtSongMen 
         Height          =   300
         Left            =   8805
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboOutDept 
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgInDetail 
         Height          =   3735
         Left            =   0
         TabIndex        =   6
         Top             =   960
         Width           =   12495
         _cx             =   22040
         _cy             =   6588
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
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
         RowHeightMax    =   0
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
         TabBehavior     =   1
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
      Begin MSComCtl2.DTPicker dtpOuterDate 
         Height          =   300
         Left            =   9120
         TabIndex        =   7
         Top             =   4920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483643
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   88145923
         CurrentDate     =   39799
      End
      Begin VB.Label lblOuter 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "接收人"
         Height          =   180
         Left            =   6360
         TabIndex        =   21
         Top             =   4920
         Width           =   540
      End
      Begin VB.Label lblApplyDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运送人"
         Height          =   180
         Left            =   8130
         TabIndex        =   10
         Top             =   645
         Width           =   540
      End
      Begin VB.Label lblPurveyDept 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "出院科室"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblOuterDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "接收时间"
         Height          =   180
         Left            =   8280
         TabIndex        =   8
         Top             =   4920
         Width           =   720
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "病案接收编辑"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   11295
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   6240
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm病案接收编辑.frx":6C94
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15108
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
End
Attribute VB_Name = "frm病案接收编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
'Private mstrNo As String                    '具体的单据号;
Private mintEditState As Integer            '1.新增；2、修改；3、查看；
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnChange As Boolean
Private mlngApplyId As Long                  '科室ID
Private mstrPatientSum As String
Private mstrPrivs As String
Private mstrOldName As String
Private mlngCount As Long
Private mdtLend As Date
Private mblnInTo As Boolean
Private mstrDeptName As String
Private mintDblick As Integer
Private mlngModule  As Long

Public Sub ShowCard(frmMain As Form, ByVal intEditState As Integer, ByVal strPatientSum As String, Optional lngDeptId As Long = 0, Optional blnSuccess As Boolean = False, Optional ByVal lngModule As Long = 201)
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--功能:显示和编辑卡片
    '--参数:frmMain-父窗口
    '       intEditState-编辑状态
    '       lngDeptId -出院科室ID
    '--出参:blnSuccess-保存成功,true,否则false
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Set mfrmMain = frmMain
    mintEditState = intEditState
    mlngApplyId = lngDeptId
    mstrPatientSum = strPatientSum
    mblnSuccess = blnSuccess
    mblnChange = False
    mlngModule = lngModule
    
    If mintEditState = 1 Then
        lblOuter.Enabled = True
        txtOuter.Enabled = True
        lblOuterDate.Enabled = True
        dtpOuterDate.Enabled = True
        With vfgInDetail
            .Editable = flexEDKbdMouse
        End With
        txtInput.Enabled = True
        chkInput.Enabled = True
        chkInput.Value = 1
    ElseIf mintEditState = 2 Then
        lblOuter.Enabled = True
        txtOuter.Enabled = True
        lblOuterDate.Enabled = True
        dtpOuterDate.Enabled = True
'        With vfgInDetail
'            .Editable = flexEDKbdMouse
'        End With
        cboOutDept.Enabled = False
        txtInput.Enabled = False
        chkInput.Enabled = False
    ElseIf mintEditState = 3 Then
        txtSongMen.Enabled = False
        lblOuter.Enabled = False
        txtOuter.Enabled = False
        lblOuterDate.Enabled = False
        dtpOuterDate.Enabled = False
        cboOutDept.Enabled = False
        txtSongMen.Enabled = False
        cmdSave.Caption = "查看(&V)"
        txtInput.Enabled = False
        chkInput.Enabled = False
    End If
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
End Sub

Private Sub cboOutDept_Change()
    mblnChange = True
End Sub

Private Sub cboOutDept_Click()
    Dim lngApplyId As Long
    
    If Me.cboOutDept.ListCount = 0 Then Exit Sub
    If Me.cboOutDept.ListIndex = -1 Then Exit Sub
    If cboOutDept.ItemData(cboOutDept.ListIndex) = 1 And cboOutDept.Text = "所有部门" Then
        lngApplyId = 0
    Else
        lngApplyId = cboOutDept.ItemData(cboOutDept.ListIndex)
    End If
    
    If lngApplyId <> mlngApplyId Then
        If Not mblnInTo Then
            mlngApplyId = lngApplyId
            Exit Sub
        End If
        If ExaminData(vfgInDetail) Then
            If MsgBox("由于出院科室发生改变,是否要清除单据中的内容(否则取消改变)?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
                mlngApplyId = lngApplyId
                mblnChange = False
                cmdSave.Enabled = False
                Call LoadvfgInDetailData(mintEditState)
            Else
                cboOutDept.ItemData(cboOutDept.ListIndex) = mlngApplyId
                cboOutDept.Text = mstrDeptName
            End If
        Else
            mlngApplyId = lngApplyId
            mblnChange = False
            cmdSave.Enabled = False
            Call LoadvfgInDetailData(mintEditState)
        End If
    End If
End Sub

Private Sub cboOutDept_GotFocus()
    With cboOutDept
        .SelStart = 0
        .SelLength = Len(.Text)
        mstrDeptName = .Text
    End With
End Sub

Private Sub cboOutDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Trim(cboOutDept.Text) = "" Then
            If vfgInDetail.Enabled Then vfgInDetail.SetFocus
            Exit Sub
        End If
        cboOutDept.Text = Replace(UCase(cboOutDept.Text), "'", "")
        vRect = GetControlRect(cboOutDept.hWnd)
        
        strSQL = "" & _
        "   SELECT A.编码, A.名称, A.简码,A.id " & _
        "   FROM  部门表 A" & _
        "   Where ( TO_CHAR (A.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or A.撤档时间 is null) AND A.ID in (" & _
        "         Select B.部门ID From 部门性质说明 B" & _
        "         Where (B.工作性质='临床' or B.工作性质='护理') and (B.服务对象=2 or B.服务对象=3)) And " & _
        "         (A.名称 like [1] or A.编码 like [1] or A.简码 like [1] or A.编码||'-'||A.名称 like [1] ) " & zl_获取站点限制(True, "A") & _
        "         start with A.上级id is null connect by prior A.id=A.上级id"
        
        strTemp = Trim(cboOutDept.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        strTemp = LfPBF & strTemp & RgPbf
        
        lngHeigth = cboOutDept.Height

        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "科室选择", False, cboOutDept.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strTemp, glngUserId)
               
        If rsTemp Is Nothing Then
            If Not blnCancel Then MsgBox "没有满足条件的科室,请检查[科室信息]!", vbInformation, gstrSysName
            If cboOutDept.Enabled Then
                cboOutDept.SetFocus
                cboOutDept.SelStart = 0
                cboOutDept.Text = mstrDeptName
                cboOutDept.SelLength = Len(cboOutDept.Text)
                Exit Sub
            End If
        End If
        With rsTemp
            If UCase(TypeName(cboOutDept)) = "COMBOBOX" Then
                cboOutDept = !编码 & "-" & IIf(IsNull(!名称), "", !名称)
                mlngApplyId = !ID
                Call GetInitDept
                zlCommFun.PressKey vbKeyTab
'                If vfgInDetail.Enabled Then Me.vfgInDetail.SetFocus
            Else
                cboOutDept.SetFocus
                cboOutDept.SelStart = 0
                cboOutDept.SelLength = Len(cboOutDept.Text)
                If cboOutDept.Enabled Then cboOutDept.SetFocus
                zlCommFun.PressKey vbKeyTab
            End If
            .Close
        End With
    End If
End Sub

Private Sub cboOutDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub chkInput_Click()
    txtInput.Enabled = IIf(chkInput.Value = 1, True, False)
End Sub

Private Sub chkInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdPrint_Click()
    printbill
End Sub

Private Sub cmdSave_Click()
     '进行数据保存处理
    Dim blnSuccess As Boolean
    Dim strBillPrint As String

    If mintEditState = 3 Then '查看
        '处理打印
        Unload Me
        Exit Sub
    End If
    
    If Not ExamineMtlBeData(vfgInDetail) Then Exit Sub
    If ExamineMtlDataRepeat(vfgInDetail) Then Exit Sub
    If Not ValidData Then Exit Sub
    
            
    blnSuccess = SaveInCard
    
    If blnSuccess = True Then
'        '修改功能:保存成功:需要检查是否自动审核
'        strBillPrint = "存盘打印"
'

        If mlngModule = 201 Then
'           If IIf(Val(zlDatabase.GetPara(strBillPrint, glngSys, mlngModule)) = 1, 1, 0) = 1 Then
'               '打印
'               printbill
'           End If
        Else
            Dim lngRow As Long
            Dim lngCurRow As Long
            For lngRow = 1 To vfgInDetail.Rows - 1
                If Val(vfgInDetail.TextMatrix(lngRow, vfgInDetail.ColIndex("病人ID"))) > 0 Then
                    lngCurRow = lngCurRow + 1
                End If
            Next
            
            MsgBox "当前总共接收病案: " & lngCurRow & " 份", vbInformation, "提示"
            
            If IIf(Val(zlDatabase.GetPara("打印接收清单", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                Call zlRptPrint(1)
            End If
        End If
        
        If mintEditState = 2 Then  '修改
            Unload Me
            Exit Sub
        End If
'        stbThis.Panels(2).Text = "上一张的单据号：" & mstrNo
    Else
        Exit Sub
    End If
    txtSongMen = ""
    txtInput = ""
    mstrPatientSum = ""
    Call LoadvfgInDetailData(mintEditState)
    mblnChange = False
    cmdSave.Enabled = False
End Sub
  
Private Function SaveInCard() As Boolean
    '----------------------------------------------------------------------------
    '--功能:保存数据
    '--返回:保存成功,返回true,否则返回false
    '----------------------------------------------------------------------------
    Dim lngOutDeptId As Long
    Dim strSongMen As String
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strRecDate As String
    Dim strOuter As String
    Dim strOutDate As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim cllTemp As New Collection
    Dim strNow As String
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim bln共享 As Boolean
    
    SaveInCard = False
    
    lngOutDeptId = cboOutDept.ItemData(cboOutDept.ListIndex)
    
    strRecDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:mm:ss")
    strSongMen = Trim(txtSongMen.Text)
    strOuter = Trim(txtOuter.Text)
    If strOuter = "" Then
        strOutDate = ""
    Else
        strOutDate = Format(dtpOuterDate, "yyyy-mm-dd HH:mm:ss")
    End If
    
    If mlngModule = 201 Then
        bln共享 = (glngHIS共享号 > 0)
    Else
        bln共享 = False
    End If
    
    If strNow = "" Then strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With vfgInDetail
        For intRow = 1 To .Rows - 1
            If Trim(.TextMatrix(intRow, .ColIndex("病人ID"))) <> "" Then
                lngPatientlId = Val(.TextMatrix(intRow, .ColIndex("病人ID")))
                lngMtyId = Val(.TextMatrix(intRow, .ColIndex("主页ID")))
                'Create Or Replace Procedure Zl_病案接收记录_Insert
'                If mintEditState = 1 Then
'                    strSQL = "   Zl_病案接收记录_Insert("
'                Else
'                    strSQL = "   Zl_病案接收记录_Update("
'                End If
                '51584:刘鹏飞,2012-12-5,接收同时完成归档
                If bln共享 = True Then
                    '完成新版护士站所有护理文件归档
                    gstrSQL = "Select distinct nvl(婴儿,0) 序号 From 病人护理文件 where 病人ID=[1] And 主页ID=[2]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病人护理文件", lngPatientlId, lngMtyId)
                    Do While Not rsTemp.EOF
                        strSQL = "  ZL_病人护理文件_ARCHIVE("
                        strSQL = strSQL & "" & lngPatientlId & ","
                        strSQL = strSQL & "" & lngMtyId & ","
                        strSQL = strSQL & "" & Val(NVL(rsTemp!序号)) & ",1)"
                        AddArray cllTemp, strSQL
                    rsTemp.MoveNext
                    Loop
                    '完成老版护士站所有护理文件归档
                    gstrSQL = "Select distinct nvl(婴儿,0) 序号 From 病人护理记录 where 病人ID=[1] And 主页ID=[2] And 病人来源 = 2"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病人护理记录", lngPatientlId, lngMtyId)
                    Do While Not rsTemp.EOF
                        strSQL = "  Zl_电子护理记录_Archive("
                        strSQL = strSQL & "" & lngPatientlId & ","
                        strSQL = strSQL & "" & lngMtyId & ","
                        strSQL = strSQL & "" & Val(NVL(rsTemp!序号)) & ","
                        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                        strSQL = strSQL & "To_Date('" & strNow & "','YYYY-MM-DD hh24:mi:ss'))"
                        AddArray cllTemp, strSQL
                    rsTemp.MoveNext
                    Loop
                    '完成所有住院病历文件归档
                    gstrSQL = "select ID from 电子病历记录 where 病人ID=[1] and 主页ID=[2] And 病历种类=2 And 病人来源=2 And RowNum<2"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "电子病历记录", lngPatientlId, lngMtyId)
                    If Not rsTemp.EOF Then
                        strSQL = "   Zl_电子病历记录_Archive(" & rsTemp!ID & ",0,1)"
                        AddArray cllTemp, strSQL
                    End If
                End If
                
                strSQL = "   Zl_病案接收记录_Insert("
                strSQL = strSQL & "" & lngPatientlId & ","
                strSQL = strSQL & "" & lngMtyId & ","
                strSQL = strSQL & "" & IIf(strSongMen = "", "NULL", "'" & strSongMen & "'") & ","
                strSQL = strSQL & "" & IIf(strOuter = "", "NULL", "'" & strOuter & "'") & ","
                strSQL = strSQL & "" & IIf(strOutDate = "", "NULL", "to_date('" & strOutDate & "','yyyy-mm-dd hh24:mi:ss')") & ","
                strSQL = strSQL & "" & IIf(strRecDate = "", "NULL", "to_date('" & strRecDate & "','yyyy-mm-dd hh24:mi:ss')") & ")"
                AddArray cllTemp, strSQL
                
                If mlngModule <> 201 Then
                '如果是电子病案接收,需要提交病案提交记录
                    strSQL = "zl_病案提交记录_Receive('" & Val(.TextMatrix(intRow, .ColIndex("提交ID"))) & "','" & strOuter & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                    AddArray cllTemp, strSQL
                End If
                
            End If
        Next
    End With
    
    Err = 0: On Error GoTo errHand:
    blnTrans = True
    ExecuteProcedureArrAy cllTemp, Me.Caption
    mblnSuccess = True
    mblnChange = False
    SaveInCard = True
    Exit Function
errHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    SaveInCard = False
End Function

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, 4)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdFind_Click()
    If lblName.Visible = False Then
        lblName.Visible = True
        txtName.Visible = True
        txtName.SetFocus
    Else
        txtName.Text = Replace(txtName.Text, "'", "")
        SearchRow vfgInDetail, vfgInDetail.ColIndex("住院号"), txtName.Text, True
        lblName.Visible = False
        txtName.Visible = False
    End If
End Sub

Private Sub cmd运送人_Click()
    Call SelectDoctor
End Sub

Private Sub Form_Load()
    mlngCount = 0
    mstrPrivs = gstrPrivs
    mblnInTo = False
    mintDblick = 0
    lblTitle = GetUnitName & lblTitle
    If Not GetInitDept Then Exit Sub
    
    If mintEditState = 1 Or mintEditState = 2 Then
        Me.txtOuter = gstrUserName
        Me.dtpOuterDate = Format(zlDatabase.Currentdate, "yyyy-MM-DD HH:mm:ss")
    End If
    
    If mlngModule = 201 Then
        Call LoadvfgInDetailData(mintEditState)
        cmd运送人.Visible = False
    Else
        Call LoadvfgInDetailAuditData(mintEditState)
        cmd运送人.Visible = True
    End If
    
    Me.cmdPrint.Visible = False
'    If mintEditState >= 3 Then
''        Me.cmdPrint.Visible = InStr(1, mstrPrivs, ";单据打印;") <> 0
'    Else
'        Me.cmdPrint.Visible = False
'    End If
    mblnInTo = True
    
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Height < 7230 Then Me.Height = 7230
    If Me.Width < 11595 Then Me.Width = 11595
    
    With PicMain
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - fraCmd.Height  '- 100
    End With
    
    With lblTitle
        .Top = 120
        .Left = 0
        .Width = PicMain.Width
    End With
    
    With vfgInDetail
        .Top = 960
        .Left = 0
        .Width = PicMain.Width
        .Height = PicMain.Height - .Top - 720
    End With
    
    With txtSongMen
        .Left = vfgInDetail.Width - .Width - 420
        lblApplyDept.Left = .Left - lblApplyDept.Width - 60
        cmd运送人.Left = .Left + .Width + 60
    End With
       
    With dtpOuterDate
        .Top = vfgInDetail.Top + vfgInDetail.Height + 225
        .Left = PicMain.Width - PicMain.Left - dtpOuterDate.Width - 120
        lblOuterDate.Left = .Left - lblOuterDate.Width - 60
    End With
    
    With txtOuter
        .Top = vfgInDetail.Top + vfgInDetail.Height + 225
        .Left = lblOuterDate.Left - .Width - 120
        lblOuter.Left = .Left - lblOuter.Width - 60
    End With
    
    txtInput.Top = dtpOuterDate.Top
    
    lblOuter.Top = dtpOuterDate.Top + 60
    lblOuterDate.Top = lblOuter.Top
    chkInput.Top = lblOuter.Top
    
    With fraCmd
        .Top = PicMain.Top + PicMain.Height
        .Left = PicMain.Left
        .Width = PicMain.Width
    End With
    
    With cmdCancel
        .Left = fraCmd.Width - .Width - 375
    End With
    
    With cmdSave
        .Left = cmdCancel.Left - .Width - 105
    End With
    
    With cmdPrint
        .Left = cmdSave.Left - .Width - 105
    End With
End Sub

Private Function GetInitDept() As Boolean
    '----------------------------------------------------------------------------
    '功能:获取出院科室
    '返回:如有出院科室,则返回True,否则返回False
    '----------------------------------------------------------------------------
    Dim strSQL As String
    Dim i As Long
    Dim blnHaving As Boolean
    Dim rsApplys As New ADODB.Recordset
    
    strSQL = "" & _
    "   SELECT A.编码, A.名称, A.简码,A.id " & _
    "   FROM  部门表 A" & _
    "   Where ( TO_CHAR (A.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or A.撤档时间 is null) AND A.ID in (" & _
    "         Select B.部门ID From 部门性质说明 B" & _
    "         Where (B.工作性质='临床' or B.工作性质='护理') and (B.服务对象=2 or B.服务对象=3)) " & zl_获取站点限制(True, "A") & _
    "         start with A.上级id is null connect by prior A.id=A.上级id"
    
    On Error GoTo errHandle
    Set rsApplys = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With rsApplys

        If .EOF Then
            GetInitDept = False
            Exit Function
        End If
    End With
    With Me.cboOutDept
        .Clear
        
        .AddItem "所有科室"
        '装入数据
        blnHaving = False
        mlngCount = rsApplys.RecordCount
        For i = 1 To rsApplys.RecordCount
            .AddItem rsApplys!编码 & "-" & rsApplys!名称
            .ItemData(.NewIndex) = rsApplys!ID
            If rsApplys!ID = mlngApplyId Then
                .ListIndex = .NewIndex
                blnHaving = True
            End If
            rsApplys.MoveNext
        Next
        rsApplys.Close
        If Not blnHaving Then
            .ListIndex = 0
        End If
    End With
    GetInitDept = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initVfgInHeadTitle()
    Dim strHead As String
    strHead = "序号,600,1,1;住院号,1500,1,1;姓名,900,1,0;性别,500,4,0;年龄,500,7,0;住院次数,900,7,0;入院科室,1200,1,0;入院时间,1100,1,0;" & _
              "出院科室,1200,1,0;出院时间,1100,1,0;出生日期,1100,1,0;家庭地址,1350,1,0;病人ID,0,7,-1;主页id,0,7,-1;提交id,0,7,-1"
    Call SetVsFlexGridChangeHead(strHead, vfgInDetail, 1)
End Sub

Private Sub LoadvfgInDetailData(ByVal intEditState As Long)
    '获取申请科室数据
    Dim strSQL As String
    Dim strBillHead As String
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngApplyId As Long
    Dim lngGeneralId As Long
    Dim i As Long
    ' " From 病案主页 U, 病人信息 X,病案接收记录 A,Table(Cast(f_Str2list('" & mstrPatientSum & "') As zlTools.t_Strlist)) B" & _
    '
    On Error GoTo errHandle
    If mintEditState = 1 Then
        strBillHead = " " & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期" & _
        " From 病案主页 U, 病人信息 X,Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) B" & _
        " Where U.病人id = X.病人id And U.主页ID <> 0 And U.出院科室id =[1] And U.病人id = substr(B.Column_Value,1,instr(B.Column_Value,'_')-1) And " & _
        "       U.主页ID = substr(B.Column_Value,instr(B.Column_Value,'_')+1)"
        strSQL = "" & _
        "   Select distinct A.病人id,A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
        "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
        "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室," & _
        "    A.住院天数, A.费用和, A.是否随诊,A.编目员, A.编目日期 " & _
        "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
        "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
        "    Order by A.出院日期 desc "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyId, mstrPatientSum)
    Else
        strBillHead = " " & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间," & _
        "      Decode(Nvl(to_char(U.编目日期), '0'), '0', '已接收', '已编目') As 状态" & _
        " From 病案主页 U, 病人信息 X,病案接收记录 A,Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) B" & _
        " Where U.病人id = X.病人id And U.主页ID <> 0 And U.出院科室id =[1] And A.病人id = substr(B.Column_Value,1,instr(B.Column_Value,'_')-1) And " & _
        "       A.主页ID = substr(B.Column_Value,instr(B.Column_Value,'_')+1) And A.病人id = U.病人id And A.主页ID = U.主页ID"
        strSQL = "" & _
        "   Select distinct A.病人id, A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
        "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
        "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室," & _
        "    A.住院天数, A.费用和, A.是否随诊,A.编目员, A.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间,A.状态" & _
        "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
        "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
        "    Order by A.出院日期 desc "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyId, mstrPatientSum)
        
    End If
      
    With vfgInDetail
        .Clear
        Call initVfgInHeadTitle
        .Rows = IIf(rsTemp.EOF, 0, rsTemp.RecordCount) + 1
        If Not rsTemp.EOF Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("序号")) = i
                .TextMatrix(i, .ColIndex("住院号")) = IIf(IsNull(rsTemp!住院号), 0, rsTemp!住院号)
                .TextMatrix(i, .ColIndex("姓名")) = IIf(IsNull(rsTemp!姓名), "", rsTemp!姓名)
                .TextMatrix(i, .ColIndex("性别")) = IIf(IsNull(rsTemp!性别), "", rsTemp!性别)
                .TextMatrix(i, .ColIndex("年龄")) = IIf(IsNull(rsTemp!年龄), "", rsTemp!年龄)
                .TextMatrix(i, .ColIndex("住院次数")) = IIf(IsNull(rsTemp!总住院次数), "", rsTemp!总住院次数)
                .TextMatrix(i, .ColIndex("入院科室")) = IIf(IsNull(rsTemp!入院科室), "", rsTemp!入院科室)
                .TextMatrix(i, .ColIndex("入院时间")) = IIf(IsNull(rsTemp!入院日期), "", Format(rsTemp!入院日期, "yyyy-MM-dd HH:mm:ss"))
                .TextMatrix(i, .ColIndex("出院科室")) = IIf(IsNull(rsTemp!出院科室), "", rsTemp!出院科室)
                .TextMatrix(i, .ColIndex("出院时间")) = IIf(IsNull(rsTemp!出院日期), "", Format(rsTemp!出院日期, "yyyy-MM-dd HH:mm:ss"))
                .TextMatrix(i, .ColIndex("出生日期")) = IIf(IsNull(rsTemp!出生日期), "", Format(rsTemp!出生日期, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("家庭地址")) = IIf(IsNull(rsTemp!家庭地址), "", rsTemp!家庭地址)
                .TextMatrix(i, .ColIndex("病人ID")) = IIf(IsNull(rsTemp!病人ID), 0, rsTemp!病人ID)
                .TextMatrix(i, .ColIndex("主页id")) = IIf(IsNull(rsTemp!主页ID), 0, rsTemp!主页ID)
                If mintEditState <> 1 Then
                    txtSongMen = IIf(IsNull(rsTemp!运送人), "", rsTemp!运送人)
                    txtOuter = IIf(IsNull(rsTemp!接收人), "", rsTemp!接收人)
                    dtpOuterDate.Value = IIf(IsNull(rsTemp!接收时间), Format(zlDatabase.Currentdate, "yyyy-MM-DD HH:mm"), Format(rsTemp!接收时间, "yyyy-MM-DD HH:mm"))
''                    lngApplyId = IIf(IsNull(rsTemp!出院科室), 0, rsTemp!出院科室)
                End If
                rsTemp.MoveNext
            Next
            If intEditState = 1 Then
                .Rows = .Rows + 1
            End If
            cmdSave.Enabled = True
        Else
            Select Case intEditState
                Case 1
                    .Rows = .Rows + 1
            End Select
        End If
        If .Rows > 1 Then
            .Select 1, .ColIndex("住院号")
        End If
        If intEditState = 1 Then
            stbThis.Panels(2).Text = "可以在‘输入住院号’输入病人住院号录入，回车增加病案接收信息！"
        End If
        .ExplorerBar = flexExSortShowAndMove
        '行选择
'        .FixedCols = 1
        .SelectionMode = flexSelectionByRow
    End With
    rsTemp.Close
    Call RestoreHead(vfgInDetail)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadvfgInDetailAuditData(ByVal intEditState As Long)
    '获取申请科室数据
    Dim strSQL As String
    Dim strBillHead As String
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngApplyId As Long
    Dim lngGeneralId As Long
    Dim i As Long
    ' " From 病案主页 U, 病人信息 X,病案接收记录 A,Table(Cast(f_Str2list('" & mstrPatientSum & "') As zlTools.t_Strlist)) B" & _
    '
    On Error GoTo errHandle
    If mintEditState = 1 Then
        strBillHead = " " & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期,A.ID as 提交ID" & _
        " From 病案主页 U, 病人信息 X,病案提交记录 A,Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) B" & _
        " Where U.病人id = X.病人id And U.主页ID <> 0 And " & IIf(mlngApplyId = 0, "0=[1]", "U.出院科室id =[1]") & " And U.病人ID = A.病人ID And U.主页ID = A.主页ID And A.记录状态=1 And U.病人id = substr(B.Column_Value,1,instr(B.Column_Value,'_')-1) And " & _
        "       U.主页ID = substr(B.Column_Value,instr(B.Column_Value,'_')+1)"
        strSQL = "" & _
        "   Select distinct A.病人id, A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
        "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
        "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室," & _
        "    A.住院天数, A.费用和, A.是否随诊,A.编目员, A.编目日期,A.提交ID " & _
        "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
        "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
        "    Order by A.出院日期 desc "
           Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyId, mstrPatientSum)
    Else
        strBillHead = " " & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间," & _
        "      Decode(Nvl(to_char(U.编目日期), '0'), '0', '已接收', '已编目') As 状态" & _
        " From 病案主页 U, 病人信息 X,病案接收记录 A,Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) B" & _
        " Where U.病人id = X.病人id And U.主页ID <> 0 And " & IIf(mlngApplyId = 0, "0=[1]", "U.出院科室id =[1]") & "  And A.病人id = substr(B.Column_Value,1,instr(B.Column_Value,'_')-1) And " & _
        "       A.主页ID = substr(B.Column_Value,instr(B.Column_Value,'_')+1) And A.病人id = U.病人id And A.主页ID = U.主页ID"
        strSQL = "" & _
        "   Select distinct A.病人id, A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
        "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
        "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室," & _
        "    A.住院天数, A.费用和, A.是否随诊,A.编目员, A.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间,A.状态" & _
        "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
        "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
        "    Order by A.出院日期 desc "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyId, mstrPatientSum)
        
    End If
      
    With vfgInDetail
        .Clear
        Call initVfgInHeadTitle
        .Rows = IIf(rsTemp.EOF, 0, rsTemp.RecordCount) + 1
        If Not rsTemp.EOF Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("序号")) = i
                .TextMatrix(i, .ColIndex("住院号")) = IIf(IsNull(rsTemp!住院号), 0, rsTemp!住院号)
                .TextMatrix(i, .ColIndex("姓名")) = IIf(IsNull(rsTemp!姓名), "", rsTemp!姓名)
                .TextMatrix(i, .ColIndex("性别")) = IIf(IsNull(rsTemp!性别), "", rsTemp!性别)
                .TextMatrix(i, .ColIndex("年龄")) = IIf(IsNull(rsTemp!年龄), "", rsTemp!年龄)
                .TextMatrix(i, .ColIndex("住院次数")) = IIf(IsNull(rsTemp!总住院次数), "", rsTemp!总住院次数)
                .TextMatrix(i, .ColIndex("入院科室")) = IIf(IsNull(rsTemp!入院科室), "", rsTemp!入院科室)
                .TextMatrix(i, .ColIndex("入院时间")) = IIf(IsNull(rsTemp!入院日期), "", Format(rsTemp!入院日期, "yyyy-MM-dd HH:mm:ss"))
                .TextMatrix(i, .ColIndex("出院科室")) = IIf(IsNull(rsTemp!出院科室), "", rsTemp!出院科室)
                .TextMatrix(i, .ColIndex("出院时间")) = IIf(IsNull(rsTemp!出院日期), "", Format(rsTemp!出院日期, "yyyy-MM-dd HH:mm:ss"))
                .TextMatrix(i, .ColIndex("出生日期")) = IIf(IsNull(rsTemp!出生日期), "", Format(rsTemp!出生日期, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("家庭地址")) = IIf(IsNull(rsTemp!家庭地址), "", rsTemp!家庭地址)
                .TextMatrix(i, .ColIndex("病人ID")) = IIf(IsNull(rsTemp!病人ID), 0, rsTemp!病人ID)
                .TextMatrix(i, .ColIndex("主页id")) = IIf(IsNull(rsTemp!主页ID), 0, rsTemp!主页ID)
                .TextMatrix(i, .ColIndex("提交id")) = IIf(IsNull(rsTemp!提交Id), 0, rsTemp!提交Id)

                If mintEditState <> 1 Then
                    txtSongMen = IIf(IsNull(rsTemp!运送人), "", rsTemp!运送人)
                    txtOuter = IIf(IsNull(rsTemp!接收人), "", rsTemp!接收人)
                    dtpOuterDate.Value = IIf(IsNull(rsTemp!接收时间), Format(zlDatabase.Currentdate, "yyyy-MM-DD HH:mm"), Format(rsTemp!接收时间, "yyyy-MM-DD HH:mm"))
''                    lngApplyId = IIf(IsNull(rsTemp!出院科室), 0, rsTemp!出院科室)
                End If
                rsTemp.MoveNext
            Next
            If intEditState = 1 Then
                .Rows = .Rows + 1
            End If
            cmdSave.Enabled = True
        Else
            Select Case intEditState
                Case 1
                    .Rows = .Rows + 1
            End Select
        End If
        If .Rows > 1 Then
            .Select 1, .ColIndex("住院号")
        End If
        If intEditState = 1 Then
            stbThis.Panels(2).Text = "可以在‘输入住院号’输入病人住院号录入，回车增加病案接收信息！"
        End If
        .ExplorerBar = flexExSortShowAndMove
        '行选择
'        .FixedCols = 1
        .SelectionMode = flexSelectionByRow
    End With
    rsTemp.Close
    Call RestoreHead(vfgInDetail)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not mblnChange Or mintEditState = 2 Or mintEditState = 4 Or mintEditState = 3 Then
        Call SaveHead(vfgInDetail)
        SaveWinState Me, App.ProductName
        Exit Sub
    End If
    
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        Call SaveHead(vfgInDetail)
        SaveWinState Me, App.ProductName
    End If
End Sub

Private Function ValidData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证数据的有效性
    '返回:验证成功返回true,否则false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ValidData = False
    Dim intLop As Integer
    
    If Trim(txtSongMen.Text) = "" Then
        MsgBox "运送人必须输入!", vbInformation + vbOKOnly, gstrSysName
        If txtSongMen.Enabled Then txtSongMen.SetFocus
        Exit Function
    End If
    
    If InStr(1, txtSongMen.Text, "'") > 0 Then
        MsgBox "运送人存在非法字符!", vbInformation + vbOKOnly, gstrSysName
        If txtSongMen.Enabled Then txtSongMen.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(Trim(txtSongMen.Text)) > 20 Then
        MsgBox "运送人超长,最多能输入10个汉字或20个字符!", vbInformation + vbOKOnly, gstrSysName
        If txtSongMen.Enabled Then txtSongMen.SetFocus
        Exit Function
    End If
    
    If Trim(txtOuter.Text) = "" Then
        MsgBox "接收人必须输入!", vbInformation + vbOKOnly, gstrSysName
        If txtOuter.Enabled Then txtOuter.SetFocus
        Exit Function
    End If
    
    If InStr(1, txtOuter.Text, "'") > 0 Then
        MsgBox "接收人存在非法字符!", vbInformation + vbOKOnly, gstrSysName
        If txtOuter.Enabled Then txtOuter.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(Trim(txtOuter.Text)) > 20 Then
        MsgBox "接收人超长,最多能输入10个汉字或20个字符!", vbInformation + vbOKOnly, gstrSysName
        If txtOuter.Enabled Then txtOuter.SetFocus
        Exit Function
    End If
    
    ValidData = True
End Function

Private Sub txtInput_GotFocus()
    With txtInput
        .SelStart = 0
        .SelLength = Len(.Text)
        mstrDeptName = .Text
    End With
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strBillHead As String
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    Dim lngPatientId As Long
    Dim lngMtalId As Long
    Dim i As Long
    Dim j As Long
    Dim strMsg As String
    Dim strSection As String
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtInput.Text) = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        txtInput.Text = Replace(UCase(txtInput.Text), "'", "")
        vRect = GetControlRect(txtInput.hWnd)
        
'        Rownum as ID
        If cboOutDept.Text = "所有科室" Then
            strSection = ""
        Else
            strSection = " And U.出院科室id =[1] "
        End If
        
        If mlngModule = 201 Then
            strBillHead = " " & _
            " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
            "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
            "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
            "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊" & _
            " From 病案主页 U, 病人信息 X" & _
            " Where U.病人id = X.病人id And U.病人性质 = 0 And U.主页ID <> 0 And U.编目日期 is null And U.出院日期 is not null " & strSection & " And U.住院号 = [2]"
            '62940:刘鹏飞,2013-06-24,已接收就不能再次接收
            If mintEditState = 1 Then
                strBillHead = strBillHead & _
                    " And NOT Exists (Select ID From 病案接收记录 Where 病人ID=U.病人ID And 主页ID=U.主页ID)"
            Else
                strBillHead = strBillHead & _
                    " And Exists (Select ID From 病案接收记录 Where 病人ID=U.病人ID And 主页ID=U.主页ID)"
            End If
            strSQL = "" & _
            "   Select distinct  Rownum as ID,A.病人id, A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
            "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
            "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室," & _
            "    A.住院天数, A.费用和, A.是否随诊,A.入院科室id, A.出院科室id " & _
            "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
            "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
            "    Order by A.出院日期 desc "
        Else
            strBillHead = " " & _
            " Select Distinct X.病人id, U.主页id, C.ID as 提交ID,U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
            "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
            "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id,U.病案状态,Decode(Nvl(U.病案状态,1),1,'提交待收',10,'接收待审',2,'拒绝接收',3,'正在审查',4,'审查反馈',5,'审查归档',6,'审查整改',13,'正在抽查',14,'抽查反馈',16,'抽查整改') as 病案状态值," & _
            "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊" & _
            " From 病案主页 U, 病人信息 X,病案提交记录 C" & _
            " Where U.病人id = X.病人id And U.病人id = C.病人ID And U.主页ID = C.主页ID And  C.记录状态 <>2 And U.主页ID <> 0   " & strSection & "  And U.住院号 = [2]"
            strSQL = "" & _
            "   Select distinct  Rownum as ID,A.病人id, A.主页id,A.提交ID, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
            "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
            "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室," & _
            "    A.住院天数, A.费用和, A.是否随诊,A.入院科室id, A.出院科室id,A.病案状态,A.病案状态值 " & _
            "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
            "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
            "    Order by A.出院日期 desc "
        End If
        
            
        strTemp = Trim(txtInput.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
'        strTemp = LfPBF & strTemp & RgPbf
        strTemp = strTemp
        lngHeigth = txtInput.Height
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyId, strTemp)
        If rsTemp.RecordCount <> 1 Then
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病案选择", False, txtInput.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, mlngApplyId, strTemp)
        End If
        
        If mlngModule <> 201 Then
            
            If rsTemp Is Nothing Then
                MsgBox "当前病案还未提交或当前指定科室没有该病案,请检查信息!", vbInformation, gstrSysName
                If txtInput.Enabled Then
                    txtInput.SetFocus
                    txtInput.SelStart = 0
                    txtInput.SelLength = Len(txtInput.Text)
                    
                    Exit Sub
                End If
            End If
            
            If rsTemp.RecordCount = 1 Then

                If rsTemp!病案状态 <> 1 Then
                    strMsg = "当前病案状态为:[" & rsTemp!病案状态值 & "],不能在进行接收!" & vbCrLf & vbCrLf
                    strMsg = strMsg & ChkStrUniCode("姓名：" & IIf(IsNull(rsTemp!姓名), "", rsTemp!姓名) & "                    ", 20) & "住院号：" & IIf(IsNull(rsTemp!住院号), 0, rsTemp!住院号)

                    MsgBox strMsg, vbInformation, gstrSysName
                    If txtInput.Enabled Then
                        txtInput.SetFocus
                        txtInput.SelStart = 0
                        txtInput.SelLength = Len(txtInput.Text)
                    End If

                    Exit Sub
                End If
            End If
        Else
        
            If rsTemp Is Nothing Then
                MsgBox "没有满足条件的病案,请检查信息!", vbInformation, gstrSysName
                If txtInput.Enabled Then
                    txtInput.SetFocus
                    txtInput.SelStart = 0
                    txtInput.SelLength = Len(txtInput.Text)
                    Exit Sub
                End If
            End If
        End If

        
        With rsTemp
            If UCase(TypeName(txtInput)) = "TEXTBOX" Then
                lngPatientId = IIf(IsNull(!病人ID), 0, !病人ID)
                lngMtalId = IIf(IsNull(!主页ID), 0, !主页ID)
                If Not ExamineInputRepeat(vfgInDetail, lngPatientId, lngMtalId) Then
                    i = 0
                    For j = 1 To vfgInDetail.Rows - 1
                        If IsNull(vfgInDetail.TextMatrix(j, vfgInDetail.ColIndex("姓名"))) Or vfgInDetail.TextMatrix(j, vfgInDetail.ColIndex("姓名")) = "" Then
                            i = j
                            j = vfgInDetail.Rows
                        End If
                    Next
                    If i = 0 Then
                        vfgInDetail.Rows = vfgInDetail.Rows + 1
                        vfgInDetail.Row = vfgInDetail.Rows - 1
                    End If
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("序号")) = vfgInDetail.Rows - 1
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("住院号")) = IIf(IsNull(!住院号), 0, !住院号)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("姓名")) = IIf(IsNull(!姓名), "", !姓名)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("性别")) = IIf(IsNull(!性别), "", !性别)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("年龄")) = IIf(IsNull(!年龄), "", !年龄)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("住院次数")) = IIf(IsNull(!总住院次数), "", !总住院次数)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("入院科室")) = IIf(IsNull(!入院科室), "", !入院科室)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("入院时间")) = IIf(IsNull(!入院日期), "", Format(!入院日期, "yyyy-MM-dd HH:mm:ss"))
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("出院科室")) = IIf(IsNull(!出院科室), "", !出院科室)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("出院时间")) = IIf(IsNull(!出院日期), "", Format(!出院日期, "yyyy-MM-dd HH:mm:ss"))
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("出生日期")) = IIf(IsNull(!出生日期), "", Format(!出生日期, "yyyy-MM-dd"))
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("家庭地址")) = IIf(IsNull(!家庭地址), "", !家庭地址)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("病人ID")) = IIf(IsNull(!病人ID), 0, !病人ID)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("主页id")) = IIf(IsNull(!主页ID), 0, !主页ID)
                    If mlngModule = 201 Then
                        vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("提交id")) = 0
                    Else
                        vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("提交id")) = IIf(IsNull(!提交Id), 0, !提交Id)
                    End If
                    
                    vfgInDetail.Rows = vfgInDetail.Rows + 1
                    vfgInDetail.Select i, vfgInDetail.ColIndex("住院号")
                    vfgInDetail.Row = i
                    mblnChange = True
                    cmdSave.Enabled = True
                    
                    txtInput.SetFocus
                    txtInput.SelStart = 0
                    txtInput.SelLength = Len(txtInput.Text)
                    If txtInput.Enabled Then txtInput.SetFocus
                    
                    
                Else
                    txtInput.SetFocus
                    txtInput.SelStart = 0
                    txtInput.SelLength = Len(txtInput.Text)
                    If txtInput.Enabled Then txtInput.SetFocus
                End If
            Else
                txtInput.SetFocus
                txtInput.SelStart = 0
                txtInput.SelLength = Len(txtInput.Text)
                If txtInput.Enabled Then txtInput.SetFocus
                zlCommFun.PressKey vbKeyTab
            End If
            
            
            
            
            
            .Close
        End With
        
        Call ShowReceiveNum
    
    End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtInput, KeyAscii, m数字式
'    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
'        Exit Sub
'    Else
'        If KeyAscii = vbKeyReturn Then
'        Else
'            KeyAscii = 0
'        End If
'    End If
End Sub

Private Sub txtOuter_GotFocus()
    txtOuter.SelStart = 0
    txtOuter.SelLength = Len(txtOuter)
End Sub

Private Sub txtOuter_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtOuter.Text) = "" Then
            txtOuter.Text = gstrUserName
            If dtpOuterDate.Enabled Then dtpOuterDate.SetFocus
            If cmdSave.Enabled Then cmdSave.SetFocus
            Exit Sub
        End If
        txtOuter.Text = Replace(UCase(txtOuter.Text), "'", "")
        vRect = GetControlRect(txtOuter.hWnd)
        
        strSQL = "" & _
            "   Select 编号,简码,姓名,id " & _
            "   From 人员表 " & _
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) " & zl_获取站点限制(True) & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
            "   order by 编号"
            
        strTemp = Trim(txtOuter.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        strTemp = LfPBF & strTemp & RgPbf
        
        lngHeigth = txtOuter.Height

        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "人员选择", False, txtOuter.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strTemp)
               
        If rsTemp Is Nothing Then
            MsgBox "没有满足条件的姓名,请检查[人员信息]!", vbInformation, gstrSysName
            If txtOuter.Enabled Then
                txtOuter.SetFocus
                txtOuter.SelStart = 0
                txtOuter.SelLength = Len(txtOuter.Text)
                Exit Sub
            End If
        End If
       
        With rsTemp
            If UCase(TypeName(txtOuter)) = "TEXTBOX" Then
                txtOuter = IIf(IsNull(!姓名), "", !姓名)
                mblnChange = True
                If cmdSave.Enabled Then Me.cmdSave.SetFocus
            Else
                txtOuter.SetFocus
                txtOuter.SelStart = 0
                txtOuter.SelLength = Len(txtOuter.Text)
                If txtOuter.Enabled Then txtOuter.SetFocus
                zlCommFun.PressKey vbKeyTab
            End If
            .Close
        End With
    End If
End Sub

Private Sub txtOuter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtSongMen_GotFocus()
    With txtSongMen
        .SelStart = 0
        .SelLength = Len(.Text)
        mstrDeptName = .Text
    End With
End Sub

Private Sub txtSongMen_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtSongMen.Text) = "" Then
            txtSongMen.Text = gstrUserName
'            If dtpOuterDate.Enabled Then dtpOuterDate.SetFocus
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        txtSongMen.Text = Replace(UCase(txtSongMen.Text), "'", "")
        vRect = GetControlRect(txtSongMen.hWnd)
        
        strSQL = "" & _
            "   Select 编号,简码,姓名,id " & _
            "   From 人员表 " & _
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) " & zl_获取站点限制(True) & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
            "   order by 编号"
            
        strTemp = Trim(txtSongMen.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        strTemp = LfPBF & strTemp & RgPbf
        
        lngHeigth = txtSongMen.Height

        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "人员选择", False, txtSongMen.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strTemp)
               
        If rsTemp Is Nothing Then
'            MsgBox "没有满足条件的姓名,请检查[人员信息]!", vbInformation, gstrSysName
            If txtSongMen.Enabled Then
                txtSongMen.SetFocus
                txtSongMen.SelStart = 0
                txtSongMen.SelLength = Len(txtSongMen.Text)
                Exit Sub
            End If
        End If
       
        With rsTemp
            If UCase(TypeName(txtOuter)) = "TEXTBOX" Then
                txtSongMen = IIf(IsNull(!姓名), "", !姓名)
                mblnChange = True
                zlCommFun.PressKey vbKeyTab
            Else
                txtSongMen.SetFocus
                txtSongMen.SelStart = 0
                txtSongMen.SelLength = Len(txtSongMen.Text)
                If txtSongMen.Enabled Then txtSongMen.SetFocus
                zlCommFun.PressKey vbKeyTab
            End If
            .Close
        End With
    End If
End Sub

Private Sub txtSongMen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub vfgInDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    With vfgInDetail
        Select Case Col
           Case vfgInDetail.ColIndex("住院号")
                strValue = Trim(.TextMatrix(Row, .ColIndex("住院号")))
                If Not IsNull(strValue) Then
                    If Not GetSelectMuchPurvey(vfgInDetail, strValue, 1) Then
                        .Select Row, Col
                        .TextMatrix(Row, .ColIndex("住院号")) = mstrOldName
                        mstrOldName = ""
                        Exit Sub
                    End If
                End If

        End Select
    End With
End Sub

Private Sub vfgInDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfgInDetail.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    If mintEditState = 1 Then
        Select Case Col
            Case vfgInDetail.ColIndex("住院号")
                mstrOldName = Trim(vfgInDetail.TextMatrix(Row, vfgInDetail.ColIndex("住院号")))
                Cancel = False
                Exit Sub
            Case Else
                Cancel = True
                Exit Sub
        End Select
    End If
End Sub

Private Sub vfgInDetail_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Select Case Col
        Case vfgInDetail.ColIndex("住院号")
            Position = -1
            Exit Sub
    End Select
    If Position = 0 Then
        Position = Col
    End If
End Sub

Private Sub vfgInDetail_BeforeSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridBeforeSort(vfgInDetail, Col, Order)
End Sub

Private Sub vfgInDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As Recordset
    Dim i As Long
    Dim j As Long
    Dim strTemp As String
    Dim lngRow As Long
    Dim lngCurrRow As Long
    Dim blnRow As Boolean
    Dim strValue As String
    
    strValue = ""
    If InStr(vfgInDetail.Cell(flexcpText, 0, Col), "住院号") > 0 Then ' And mintDblick = 0
         Err = 0: On Error GoTo errHand:
        If Not GetSelectMuchPurvey(vfgInDetail, strValue, 2) Then
            vfgInDetail.Select Row, Col
            Exit Sub
        End If
        Exit Sub
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
'

Private Sub vfgInDetail_Click()
    mintDblick = 0
End Sub

Private Sub vfgInDetail_DblClick()
    mintDblick = 1
End Sub

Private Sub vfgInDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If mintEditState = 1 Or mintEditState = 2 Then
        If KeyCode = vbKeyDelete Then
            If vfgInDetail.Row > 0 Then
                If MsgBox("真要删除当前记录吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    vfgInDetail.RemoveItem (vfgInDetail.Row)
                    If vfgInDetail.Row = 0 Then
                        vfgInDetail.Rows = vfgInDetail.Rows + 1
                        vfgInDetail.Select vfgInDetail.Rows - 1, vfgInDetail.Col
                    End If
                End If
            End If
            
            Call ShowReceiveNum
        End If
        
        If KeyCode = vbKeyInsert Then
            With vfgInDetail
                If MsgBox("真要增加记录吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                   vfgInDetail.Rows = vfgInDetail.Rows + 1
                   .Select vfgInDetail.Rows - 1, vfgInDetail.Col
                End If
                
                If mintEditState = 1 Then
                    stbThis.Panels(2).Text = "可以在‘输入住院号’输入病人住院号录入，回车增加病案接收信息！" & " 当前接收病案: " & vfgInDetail.Rows - 2 & " 份"
                End If
            End With
        End If
    End If
    If KeyCode = vbKeyReturn Then
        If mintEditState = 1 Then
            Call zlPvVsMoveGridCell(vfgInDetail, vfgInDetail.ColIndex("住院号"), vfgInDetail.ColIndex("家庭地址"), True, lngRow, SetHeadCodeData(vfgInDetail))
        Else
            Call zlPvVsMoveGridCell(vfgInDetail, vfgInDetail.ColIndex("住院号"), vfgInDetail.ColIndex("家庭地址"), False, lngRow, SetHeadCodeData(vfgInDetail))
        End If
    End If
    
    If KeyCode <> vbKeyReturn Then
        vfgInDetail.ColComboList(vfgInDetail.ColIndex("住院号")) = ""
    End If
End Sub

Private Sub vfgInDetail_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngRow As Long
    If KeyCode = vbKeyReturn Then
        If mintEditState = 1 Then
            Call zlPvVsMoveGridCell(vfgInDetail, vfgInDetail.ColIndex("住院号"), vfgInDetail.ColIndex("家庭地址"), True, lngRow, SetHeadCodeData(vfgInDetail))
        Else
            Call zlPvVsMoveGridCell(vfgInDetail, vfgInDetail.ColIndex("住院号"), vfgInDetail.ColIndex("家庭地址"), False, lngRow, SetHeadCodeData(vfgInDetail))
        End If
    End If
End Sub

Private Sub vfgInDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0 'Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub vfgInDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     Select Case KeyAscii
        Case vbKeyReturn
        Case vbKeyBack, vbKeyEscape, 3, 22: Exit Sub
        Case Else
'            Select Case Col
'                Case vfgInDetail.ColIndex("数量")
'                    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
'                        Exit Sub
'                    Else
'                        KeyAscii = 0
'                    End If
''            End Select
''            Select Case Col
'                Case vfgInDetail.ColIndex("冲销数量")
'                    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
'                        Exit Sub
'                    Else
'                        KeyAscii = 0
'                    End If
''            End Select
''            Select Case Col
'                Case vfgInDetail.ColIndex("清洗单价")
'                    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
'                        Exit Sub
'                    Else
'                        KeyAscii = 0
'                    End If
'            End Select
    End Select
End Sub

Private Sub vfgInDetail_KeyUp(KeyCode As Integer, Shift As Integer)
     vfgInDetail.ColComboList(vfgInDetail.ColIndex("住院号")) = "..."
End Sub

Private Sub vfgInDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     vfgInDetail.ColComboList(vfgInDetail.ColIndex("住院号")) = "..."
End Sub

Private Function GetSelectMuchPurvey(ByRef vsGrid As VSFlexGrid, ByVal strSearch As String, ByVal intFlag As Integer) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能:根据条件,请数据
    '参数:strSearch-搜索条件值,
    '返回:当只满足一个值时返回True,否则返回False
    '--------------------------------------------------------------------------------------------------------------
    Dim LfPBF As String
    Dim RgPbf As String
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strBillHead As String
    Dim strTemp As String
    Dim lngHeigth As Long
    Dim lngTop As Long
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim StrCodeName As String
    Dim lngPatientId As Long
    Dim lngMtalId As Long
    
    Dim lngRow As Long
    Dim i As Long
    Dim j As Long
    Dim strSection As String
    
    If intFlag = 1 Then
        If strSearch = "" Then Exit Function
        If InStr(1, strSearch, "'") <> 0 Then
            MsgBox "输入值中含有非法字符！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        
        strSearch = Replace(UCase(strSearch), "'", "")
        
'        strTemp = LfPBF & strSearch & RgPbf
        strTemp = strSearch
    End If
    
    If cboOutDept.Text = "所有科室" Then
        strSection = ""
    Else
        strSection = " And U.出院科室id =[1] "
    End If
        
    vRect = GetControlRect(vsGrid.hWnd)
    '62940:刘鹏飞,2013-06-24
    If mlngModule = 201 Then
        If intFlag = 1 Then
            strBillHead = " " & _
            " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
            "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
            "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
            "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊" & _
            " From 病案主页 U, 病人信息 X" & _
            " Where U.病人id = X.病人id And U.病人性质 = 0 And U.主页ID <> 0 And U.编目日期 is null And U.出院日期 is not null " & strSection & " And U.住院号 = [2]"
        Else
            strBillHead = " " & _
            " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
            "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
            "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
            "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊" & _
            " From 病案主页 U, 病人信息 X" & _
            " Where U.病人id = X.病人id And U.病人性质 = 0 And U.主页ID <> 0 And U.编目日期 is null And U.出院日期 is not null " & strSection
        End If
        '62940:刘鹏飞,2013-06-24,已接收就不能再次接收
        If mintEditState = 1 Then
            strBillHead = strBillHead & _
            " And NOT Exists (Select ID From 病案接收记录 Where 病人ID=U.病人ID And 主页ID=U.主页ID)"
        Else
            strBillHead = strBillHead & _
            " And  Exists (Select ID From 病案接收记录 Where 病人ID=U.病人ID And 主页ID=U.主页ID)"
        End If
        strSQL = "" & _
            "   Select distinct  Rownum as ID,A.病人id, A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
            "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
            "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室," & _
            "    A.住院天数, A.费用和, A.是否随诊,A.入院科室id, A.出院科室id " & _
            "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
            "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
            "    Order by A.出院日期 desc "
    Else
        If intFlag = 1 Then
            strBillHead = " " & _
            " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
            "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
            "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
            "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊,A.ID as 提交ID" & _
            " From 病案主页 U, 病人信息 X,病案提交记录 A" & _
            " Where U.病人id = X.病人id And U.病人ID = A.病人ID And U.主页ID = A.主页ID And A.记录状态=1 And U.主页ID <> 0 And U.编目日期 is null " & strSection & " And U.住院号 = [2]"
        Else
            strBillHead = " " & _
            " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
            "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
            "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
            "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊,A.ID as 提交ID" & _
            " From 病案主页 U, 病人信息 X,病案提交记录 A" & _
            " Where U.病人id = X.病人id And U.主页ID <> 0 And  U.病人ID = A.病人ID And U.主页ID = A.主页ID And A.记录状态=1 And U.编目日期 is null " & strSection
        End If
        strSQL = "" & _
        "   Select distinct  Rownum as ID,A.病人id, A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
        "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
        "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室," & _
        "    A.住院天数, A.费用和, A.是否随诊,A.入院科室id, A.出院科室id,A.提交ID " & _
        "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
        "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
        "    Order by A.出院日期 desc "
    
    End If
    Err = 0
    On Error GoTo errHand:
    
    lngTop = vRect.Top + vsGrid.CellTop + vsGrid.CellHeight

    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病案选择", False, vsGrid.Tag, "", False, False, True, vRect.Left - 15, lngTop, lngHeigth, blnCancel, False, False, mlngApplyId, strTemp)
    
    If rsTemp Is Nothing Then
        If Not blnCancel Then MsgBox "没有满足条件的病案信息!", vbInformation, gstrSysName
        If vsGrid.Enabled Then
            vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("住院号")) = mstrOldName
            vsGrid.SetFocus
            GetSelectMuchPurvey = False
            Exit Function
        End If
    End If

    i = 1
    With rsTemp
        If UCase(TypeName(vsGrid)) = "VSFLEXGRID" Then
            With vsGrid
                lngPatientId = IIf(IsNull(rsTemp!病人ID), 0, rsTemp!病人ID)
                lngMtalId = IIf(IsNull(rsTemp!主页ID), 0, rsTemp!主页ID)
                If Not ExamineInputRepeat(vfgInDetail, lngPatientId, lngMtalId) Then
                    i = 0
                    For j = 1 To .Rows - 1
                        If IsNull(.TextMatrix(j, .ColIndex("姓名"))) Or .TextMatrix(j, .ColIndex("姓名")) = "" Then
                            i = j
                            j = .Rows
                        End If
                    Next
                    If i = 0 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                    .TextMatrix(i, .ColIndex("序号")) = .Rows - 1
                    .TextMatrix(i, .ColIndex("住院号")) = IIf(IsNull(rsTemp!住院号), 0, rsTemp!住院号)
                    .TextMatrix(i, .ColIndex("姓名")) = IIf(IsNull(rsTemp!姓名), "", rsTemp!姓名)
                    .TextMatrix(i, .ColIndex("性别")) = IIf(IsNull(rsTemp!性别), "", rsTemp!性别)
                    .TextMatrix(i, .ColIndex("年龄")) = IIf(IsNull(rsTemp!年龄), "", rsTemp!年龄)
                    .TextMatrix(i, .ColIndex("住院次数")) = IIf(IsNull(rsTemp!总住院次数), "", rsTemp!总住院次数)
                    .TextMatrix(i, .ColIndex("入院科室")) = IIf(IsNull(rsTemp!入院科室), "", rsTemp!入院科室)
                    .TextMatrix(i, .ColIndex("入院时间")) = IIf(IsNull(rsTemp!入院日期), "", Format(rsTemp!入院日期, "yyyy-MM-dd HH:mm:ss"))
                    .TextMatrix(i, .ColIndex("出院科室")) = IIf(IsNull(rsTemp!出院科室), "", rsTemp!出院科室)
                    .TextMatrix(i, .ColIndex("出院时间")) = IIf(IsNull(rsTemp!出院日期), "", Format(rsTemp!出院日期, "yyyy-MM-dd HH:mm:ss"))

                    .TextMatrix(i, .ColIndex("出生日期")) = IIf(IsNull(rsTemp!出生日期), "", Format(rsTemp!出生日期, "yyyy-MM-dd"))
                    .TextMatrix(i, .ColIndex("家庭地址")) = IIf(IsNull(rsTemp!家庭地址), "", rsTemp!家庭地址)
                    .TextMatrix(i, .ColIndex("病人ID")) = IIf(IsNull(rsTemp!病人ID), 0, rsTemp!病人ID)
                    .TextMatrix(i, .ColIndex("主页id")) = IIf(IsNull(rsTemp!主页ID), 0, rsTemp!主页ID)
                    If mlngModule = 201 Then
                        .TextMatrix(i, vfgInDetail.ColIndex("提交id")) = 0
                    Else
                        .TextMatrix(i, .ColIndex("提交ID")) = IIf(IsNull(rsTemp!提交Id), 0, rsTemp!提交Id)
                    End If
                   
                    
                    .Rows = .Rows + 1
                    .Select i, .ColIndex("住院号")
                    mblnChange = True
                    cmdSave.Enabled = True
                End If
            End With
            
            Call ShowReceiveNum
            
       Else
            .Close
            If vsGrid.Enabled Then vsGrid.SetFocus
            zlCommFun.PressKey vbKeyTab
        End If
        .Close
    End With
    GetSelectMuchPurvey = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RightHead(ByVal vsGrid As VSFlexGrid)

    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = GetControlRect(vsGrid.hWnd)
    lngLeft = vRect.Left + vsGrid.Left
    lngTop = vRect.Top + vsGrid.RowHeight(0) 'vsGrid.CellTop + vsGrid.CellHeight  '
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, vsGrid.RowHeight(0))
    Call SaveHead(vsGrid)
End Sub

Private Sub SaveHead(ByVal vsGrid As VSFlexGrid)
    zl_VsGrid_SaveToPara vsGrid, Me.Caption, mlngModule, "病案接收编辑列头信息", True, False
End Sub

Private Sub RestoreHead(ByVal vsGrid As VSFlexGrid)
    zl_VsGrid_FromParaRestore vsGrid, Me.Caption, mlngModule, "病案接收编辑列头信息", True, False
End Sub

Private Sub vfgInDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intGetHeight As Integer
    Dim intGetWidth As Integer
    
    intGetWidth = vfgInDetail.ColWidth(0)
    intGetHeight = vfgInDetail.RowHeight(0)
    If (Button = 2) Then
        If X < intGetWidth And Y < intGetHeight Then
            Call RightHead(vfgInDetail)
        End If
    End If
End Sub

Private Function ExamineMtlDataRepeat(ByRef vsGrid As VSFlexGrid) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证器数据的是否有重复，以及病人的出院时间是否大于接收时间
    '返回:有重复返回true,否则false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    Dim lngMeterlId As Long
    Dim strValue As String
        
    ExamineMtlDataRepeat = True
    With vsGrid
        For i = 1 To .Rows - 1
            For j = i To .Rows - 1
                If i <> j Then
                    If .TextMatrix(i, .ColIndex("病人ID")) = .TextMatrix(j, .ColIndex("病人ID")) And .TextMatrix(i, .ColIndex("主页id")) = .TextMatrix(j, .ColIndex("主页id")) Then
                        MsgBox "在第" & i & "行的病案与第" & j & "的病案相同，请删除其中一行的数据！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            Next
        Next
    End With
    
    '62939:刘鹏飞,2013-06-24,病案接收时间不能小于病人出院时间
    With vsGrid
        For i = 1 To .Rows - 1
            If IsDate(.TextMatrix(i, .ColIndex("出院时间"))) Then
                If CDate(Format(.TextMatrix(i, .ColIndex("出院时间")), "YYYY-MM-DD HH:mm:ss")) > CDate(Format(dtpOuterDate.Value, "YYYY-MM-DD HH:mm:ss")) Then
                    MsgBox "在第" & i & "行的病案出院时间大于病案接收时间,请检查！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    
    ExamineMtlDataRepeat = False
End Function

Private Function ExamineInputRepeat(ByRef vsGrid As VSFlexGrid, ByVal lngPatientId As Long, ByVal lngMtalId As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证器数据的是否有重复
    '返回:有重复返回true,否则false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
           
    ExamineInputRepeat = True
    With vsGrid
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("病人ID"))) = lngPatientId And Val(.TextMatrix(i, .ColIndex("主页id"))) = lngMtalId Then
                MsgBox "录入病案在列表中第" & i & "的病案相同，请录入其它病案的数据！", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End With
    
    ExamineInputRepeat = False
End Function

Private Function ExamineMtlBeData(ByRef vsGrid As VSFlexGrid) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证是否存在数据
    '返回:有返回true,否则false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    ExamineMtlBeData = False
    With vsGrid
        For i = 1 To .Rows - 1
            If Not IsNull(.TextMatrix(i, .ColIndex("病人ID"))) And .TextMatrix(i, .ColIndex("病人ID")) <> "" Then
                ExamineMtlBeData = True
                Exit Function
            End If
        Next
    End With
    MsgBox "不存在的数据，不能进行保存！", vbInformation, gstrSysName
End Function

Private Function ExaminData(ByRef vsGrid As VSFlexGrid) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证是否存在数据
    '返回:有返回true,否则false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    ExaminData = False
    With vsGrid
        For i = 1 To .Rows - 1
            If Not IsNull(.TextMatrix(i, .ColIndex("病人ID"))) And .TextMatrix(i, .ColIndex("病人ID")) <> "" And .TextMatrix(i, .ColIndex("姓名")) <> "" Then
                ExaminData = True
                Exit Function
            End If
        Next
    End With
End Function

'打印单据
Private Sub printbill()
    Dim strGetData As String
    strGetData = GetChoiceData(vfgInDetail)
'    ReportOpen gcnOracle, glngSys, "ZL4_BILL_361_1", Me, "单据编号=" & strNo, "记录状态=" & mintRecordState, "单位=0", 2
End Sub

Private Function SetHeadCodeData(ByRef vsGrid As VSFlexGrid) As String
    Dim i As Long
    Dim strTemp As String
    
    SetHeadCodeData = ""
    With vsGrid
        For i = 0 To .Cols - 1
            If mintEditState = 1 Then
                If i = .ColIndex("住院号") Then
                    If IsNull(strTemp) Or strTemp = "" Then
                        strTemp = i & "||0"
                    Else
                        strTemp = strTemp & ";" & i & "||0"
                    End If
                End If
            End If
        Next
    End With
    SetHeadCodeData = strTemp
End Function

Private Function GetChoiceData(ByRef vsGrid As VSFlexGrid) As String
    Dim lngApplyId As String
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strStatus As String
    Dim lngRows As Long
    Dim strTemp As String
    Dim i As Long
    Dim j As Long
    Dim intCount As Integer
                    
    intCount = 0
    strTemp = ""
    GetChoiceData = ""
    With vsGrid
        If .Rows > 1 Then
            lngRows = .Rows - 1
            For i = 1 To lngRows
                lngPatientlId = Val(.TextMatrix(i, .ColIndex("病人ID")))
                lngMtyId = Val(.TextMatrix(i, .ColIndex("主页ID")))
                If Not IsNull(.TextMatrix(i, .ColIndex("病人ID"))) And .TextMatrix(i, .ColIndex("病人ID")) <> "" And .TextMatrix(i, .ColIndex("姓名")) <> "" Then
                    intCount = intCount + 1
                    If intCount > 100 Then
                        GetChoiceData = strTemp
                        MsgBox "你输入的病案数太多了，只处理前面选中的100份。", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If strTemp = "" Then
                        strTemp = lngPatientlId & "_" & lngMtyId
                    Else
                        strTemp = strTemp & "," & lngPatientlId & "_" & lngMtyId
                    End If
                End If
            Next
        End If
    End With
    GetChoiceData = strTemp
End Function

'按编码，名称，别名查找某一列
Private Function SearchRow(ByVal vfgBill As VSFlexGrid, ByVal intColIndex As Integer, _
    ByVal strColValue As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim strSQL As String
    Dim rsCode As New Recordset
    
    SearchRow = True
    With vfgBill
        If .Rows = 2 Then Exit Function
        If strColValue = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .Rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, intColIndex) <> "" Then
                strCode = .TextMatrix(intRow, intColIndex)
                If InStr(1, UCase(strCode), UCase(strColValue)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = intColIndex
                    .Select .Row, .Col
                    Exit Function
                End If
            End If
        Next
        
        On Error GoTo errHandle
        strSQL = "SELECT 住院号 " _
                 & "FROM 病案主页 " _
               & " Where upper(姓名) LIKE '" & IIf(gstrMatchMethod = "0", "%", "") & strColValue & "%' "
        Set rsCode = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsCode.EOF Then
            SearchRow = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, intColIndex) <> "" Then
                strCode = .TextMatrix(intRow, intColIndex)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!住院号)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = intColIndex
                        rsCode.Close
                        .Select .Row, .Col
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            
            End If
        Next
        rsCode.Close
    End With
    SearchRow = False
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub zlPvVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng主例 As Long = -1, Optional lng尾列 As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef lngRow As Long = -1, Optional strValue As String)
    ', Optional strHeadMove As String
    '-----------------------------------------------------------------------------------------------------------

    '功能:移动单元格的列
    '入参:blnEdit-当前正处于编辑状态,允许新增行
    '     lng主例-主列,如果<0,则主列为0列,否则为指定的列
    '     lng尾列-尾列,如果<0,则主列为.cols-1,否则为指定的列
    '出参:lngRow-如果存在插入行,则返回被插入的行号,否则返回-1
    '返回:
    '编制:刘兴洪
    '日期:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------

    Dim lngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long
    Dim lngValue As Long
    Dim arrHead As Variant
    Dim j As Long
    Dim lngColValue As Long
    
    Err = 0: On Error GoTo errHand:
    'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)

    If lng主例 <> -1 Then
        lngCol = lng主例
    Else
        lngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If lngCol = -1 Then lngCol = 0
    lngLastCol = IIf(lng尾列 < 0, vsGrid.Cols - 1, lng尾列)
    lngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = lngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        lngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                If IsNull(strValue) Or strValue = "" Then
                    arrSplit = Split(.ColData(i) & "||", "||")
                    If IsNull(arrSplit(1)) Or Trim(arrSplit(1)) = "" Then
                        lngValue = 0
                    Else
                        lngValue = Val(arrSplit(1))
                    End If
                Else
                    arrHead = Split(strValue, ";")
                    For j = 0 To UBound(arrHead)
                        lngValue = 1
                        lngColValue = Val(Split(arrHead(j), "||")(0))
                        If i = lngColValue Then
                            lngValue = Val(Split(arrHead(j), "||")(1))
                            Exit For
                        End If
                    Next
                End If
                If .ColHidden(i) Or lngValue >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = lngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    lngRow = .Row
                                End If
                            End If
                            .Col = lngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.Width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
errHand:
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim objRow1 As New zlTabAppRow
    Dim strRange As String
    Dim intCol As Long
    Dim strListTitle As String
    
    strListTitle = "病历科室运送病人病案情况"
    With vfgInDetail
        '清除选择行的颜色
''        For intCol = 0 To .Cols - 1
''            .Col = intCol
''            .CellBackColor = glngGetFocus_Font
'''            .CellForeColor = glngLostFocus_Font
''        Next
        .GridLines = flexGridInset
    End With
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strListTitle
        
    Set objRow = New zlTabAppRow

    If cboOutDept.Visible Then
        objRow.Add "出院科室:" & cboOutDept.Text
    End If
    objPrint.UnderAppRows.Add objRow
    
    Set objRow1 = New zlTabAppRow
    objRow1.Add "运送人:" & txtSongMen.Text
    objRow1.Add "接收人:" & txtOuter.Text
    objPrint.BelowAppRows.Add objRow1
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = vfgInDetail
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    With vfgInDetail
        .GridLines = flexGridNone
    End With
End Sub

Private Sub ShowReceiveNum()
'显示当前待接收病案的分数
    Dim lngRow As Long
    Dim lngCurRow As Long
    If mintEditState = 1 Then
        For lngRow = 1 To vfgInDetail.Rows - 1
            
            
            If Val(vfgInDetail.TextMatrix(lngRow, vfgInDetail.ColIndex("病人ID"))) > 0 Then
                vfgInDetail.TextMatrix(lngRow, vfgInDetail.ColIndex("序号")) = lngRow
                lngCurRow = lngCurRow + 1
            End If
        Next
    
        stbThis.Panels(2).Text = "可以在‘输入住院号’输入病人住院号录入，回车增加病案接收信息！" & " 当前等待接收病案: " & lngCurRow & " 份"
    End If
End Sub

'选择医生
Private Sub SelectDoctor(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
    Dim strSQL As String
On Error GoTo errH
    strSQL = ""
    If strShortName <> "" Then
        strSQL = strSQL & vbCrLf & "Select c.ID,c.编号,c.姓名 As 名称"
        strSQL = strSQL & vbCrLf & "From 人员表 C"
        strSQL = strSQL & vbCrLf & "Where  c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null "
        strSQL = strSQL & vbCrLf & "And (c.姓名 like '%'||[1]||'%' or 简码 like '%'||[1]||'%')"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strShortName))
        
        bytRet = ShowPubSelectTest(Me, txtSongMen, 2, "编号,1200,0,;名称,1200,0,", Me.Name & "\运送人选择", "请从下表中选择一个医生", rsTmp, rsResult, 8790, 4500, False)
    Else
        strSQL = strSQL & vbCrLf & "Select Id,上级id,0 As 末级,编码 as 编号,名称 From 部门表"
        strSQL = strSQL & vbCrLf & "Start With 上级id Is Null"
        strSQL = strSQL & vbCrLf & "Connect By Prior ID = 上级id"
        strSQL = strSQL & vbCrLf & "Union All"
        strSQL = strSQL & vbCrLf & "Select c.id,b.部门id As 上级Id,1 As 末级,c.编号,c.姓名 As 名称"
        strSQL = strSQL & vbCrLf & "From 部门人员 b,人员表 C"
        strSQL = strSQL & vbCrLf & "Where c.Id=b.人员id And b.缺省=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
      
        bytRet = ShowPubSelectTest(Me, txtSongMen, 3, "编号,1200,0,;名称,1200,0,", Me.Name & "\运送人选择", "请从下表中选择一个医生", rsTmp, rsResult, 8790, 4500, False)
 
    End If
    
    If rsResult Is Nothing Then
        txtSongMen.Text = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txtSongMen.Text = ""
    Else
        txtSongMen.Text = rsResult!名称
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub


