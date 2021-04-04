VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutDoctorAdvice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   Icon            =   "frmOutDoctorAdvice.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAdviceUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   60
      MousePointer    =   7  'Size N S
      TabIndex        =   3
      Top             =   4455
      Width           =   7275
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAppend 
      Height          =   1380
      Left            =   60
      TabIndex        =   2
      Top             =   4815
      Width           =   7275
      _cx             =   12832
      _cy             =   2434
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
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
   Begin MSComctlLib.TabStrip tabAppend 
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   4500
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   529
      Style           =   2
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "医嘱计价项目(&P)"
            Key             =   "医嘱计价项目"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "医嘱发送明细(&S)"
            Key             =   "医嘱发送明细"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "医嘱签名记录(&G)"
            Key             =   "医嘱签名记录"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4380
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   7260
      _cx             =   12806
      _cy             =   7726
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
      FormatString    =   $"frmOutDoctorAdvice.frx":000C
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList imgSign 
         Left            =   3285
         Top             =   1170
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":00A7
               Key             =   "签名"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgFlag 
         Left            =   1260
         Top             =   1140
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   8
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":03F9
               Key             =   "紧急"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":0613
               Key             =   "补录"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":0B2D
               Key             =   "未申请"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":1047
               Key             =   "已申请"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgPass 
         Left            =   2265
         Top             =   1155
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   14
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":1561
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":185B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":1B55
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":1E4F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":2149
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picFocus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   30
      ScaleHeight     =   600
      ScaleWidth      =   630
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   630
   End
End
Attribute VB_Name = "frmOutDoctorAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mfrmParent As Object
Public mstrPrivs As String
Private WithEvents mfrmEdit As Form
Attribute mfrmEdit.VB_VarHelpID = -1

'上次刷新数据时的病人信息
Private mlng病人ID As Long
Private mstr挂号单 As String
Private mint状态 As Integer '0-候诊病人,1-在诊病人,2-已诊病人
Private mlng前提ID As Long
Private mblnShowAll As Boolean

Private mbln自动皮试 As Boolean

Private mblnMoved As Boolean '当前挂号单是否已经转出
Private mvRegDate As Date '挂号单的挂号时间

'医嘱菜单索引
Private Enum Menu_Advice
    mnu新开医嘱 = 0
    mnu修改医嘱 = 1
    mnu删除医嘱 = 2
    mnu皮试结果 = 4 '-
    mnu发送医嘱 = 6 '-
    mnu作废医嘱 = 7
    mnu复制到文本 = 9 '-
End Enum

'报表菜单项索引
Private Enum Menu_Report
    mnu打印诊疗单据 = 0
End Enum

Private Enum COL医嘱清单
    '固定列
    COL_F标志 = 0
    COL_F申请 = 1
    '隐藏列
    COL_ID = 2
    COL_相关ID = COL_ID + 1
    COL_组ID = COL_ID + 2
    COL_组号 = COL_ID + 3
    COL_婴儿ID = COL_ID + 4
    COL_医嘱状态 = COL_ID + 5
    COL_诊疗类别 = COL_ID + 6
    COL_操作类型 = COL_ID + 7
    COL_毒理分类 = COL_ID + 8
    COL_标志 = COL_ID + 9
    '可见列
    COL_警示 = COL_ID + 10 'Pass
    COL_开始时间 = COL_ID + 11
    COL_医嘱内容 = COL_ID + 12
    COL_皮试 = COL_ID + 13
    COL_总量 = COL_ID + 14
    COL_单量 = COL_ID + 15
    COL_频率 = COL_ID + 16
    COL_用法 = COL_ID + 17
    COL_医生嘱托 = COL_ID + 18
    COL_执行时间 = COL_ID + 19
    COL_执行科室 = COL_ID + 20
    COL_执行性质 = COL_ID + 21
    COL_开嘱医生 = COL_ID + 22
    COL_开嘱时间 = COL_ID + 23
    COL_发送人 = COL_ID + 24
    COL_发送时间 = COL_ID + 25
    '隐藏列
    COL_单据ID = COL_ID + 26 '对应病历文件目录.ID
    COL_申请项 = COL_ID + 27 '诊疗单据是否有申请项
    COL_报告项 = COL_ID + 28 '诊疗单据是否有报告项
    COL_申请ID = COL_ID + 29 '对应病人病历记录.ID
    COL_前提ID = COL_ID + 30
    COL_签名否 = COL_ID + 31
End Enum

Private Enum COL发送清单
    cs发送号 = 0
    cs发送时间 = 1
    cs发送医嘱 = 2
    cs单据号 = 3
    cs收费项目 = 4
    cs数次 = 5
    cs计费状态 = 6
    cs执行状态 = 7
    cs执行科室 = 8
    cs发送人 = 9
    cs记录性质 = 10
End Enum

Public Function zlRefresh(lng病人ID As Long, str挂号单 As String, int状态 As Integer, varValue As Variant, Optional ByVal lng前提ID As Long = 0, Optional ByVal ifShowAll As Boolean = True) As Boolean
'功能：刷新或清除医嘱清单
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    mlng病人ID = lng病人ID
    mstr挂号单 = str挂号单
    mint状态 = int状态
    mlng前提ID = lng前提ID
    mblnShowAll = ifShowAll
        
    '挂号单是否转出,及挂号时间
    mblnMoved = False
    If lng病人ID <> 0 Then
        If mint状态 = 2 Then '根据业务情况,只判断已诊病人
            mblnMoved = MovedByNO(str挂号单, "病人挂号记录")
        End If
        strSQL = "Select 登记时间 From 病人挂号记录 Where NO=[1]"
        If mblnMoved Then strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, str挂号单)
        If Not rsTmp.EOF Then
            mvRegDate = rsTmp!登记时间
        Else
            mvRegDate = zlDatabase.Currentdate
        End If
        On Error GoTo 0
    End If
    
    If mlng病人ID = 0 Then
        '清除医嘱清单
        Call ClearAdviceData
        Call ClearAppendData
        mfrmParent.stbThis.Panels(2).Text = ""
    Else
        '显示医嘱清单
        Call LoadAdvice
        Call ShowTotalMoney
    End If
    zlRefresh = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlButtonClick(objButton As Button) As Boolean
'功能：执行医嘱按钮功能
    Select Case objButton.Key
        Case "新开"
            Call FuncAdviceAdd
        Case "修改"
            Call FuncAdviceModi
        Case "删除"
            Call FuncAdviceDel
        Case "发送"
            Call FuncAdviceSend
        Case "作废"
            Call FuncAdviceRevoke
        Case "签名"
            Call FuncAdviceSign
    End Select
End Function

Public Function zlMenuClick(objMenu As Menu) As Boolean
'功能：执行医嘱菜单功能
    Dim strText As String
    
    If objMenu.Caption Like "*(&*)*" Then
        strText = Split(objMenu.Caption, "(")(0)
    Else
        strText = objMenu.Caption
    End If
        
    If objMenu.Name = "mnuReportClinic" Then
        '打印诊疗单据
        Call FuncBillPrint(objMenu)
    ElseIf objMenu.Name = "mnuViewAdviceAppend" Then
        '显示/隐藏附加表格
        objMenu.Checked = Not objMenu.Checked
        fraAdviceUD.Visible = objMenu.Checked
        tabAppend.Visible = objMenu.Checked
        vsAppend.Visible = objMenu.Checked
        Call Form_Resize
        
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Else
        Select Case strText
            Case "新开医嘱"
                Call FuncAdviceAdd
            Case "修改医嘱"
                Call FuncAdviceModi
            Case "删除医嘱"
                Call FuncAdviceDel
            Case "皮试结果"
                Call FuncAdviceTest
            Case "发送医嘱"
                Call FuncAdviceSend
            Case "作废医嘱"
                Call FuncAdviceRevoke
            Case "复制到文本"
                Call FuncCopyToText
            Case "电子签名"
                Call FuncAdviceSign
            Case "取消签名"
                Call FuncAdviceSignErase
            Case "验证签名"
                Call FuncAdviceSignVerify
        End Select
    End If
End Function

Private Sub FuncCopyToText()
    Dim strCopy As String, intRow As Integer
    strCopy = ""
    With vsAdvice
        For intRow = .FixedRows To .Rows - 1
            If .TextMatrix(intRow, COL_诊疗类别) = "5" Or .TextMatrix(intRow, COL_诊疗类别) = "6" Then
                strCopy = strCopy & .TextMatrix(intRow, COL_医嘱内容) _
                        & " " & .TextMatrix(intRow, COL_单量) _
                        & " " & .TextMatrix(intRow, COL_频率) _
                        & " " & .TextMatrix(intRow, COL_用法) _
                        & vbCrLf
            Else
                strCopy = strCopy & .TextMatrix(intRow, COL_医嘱内容) & vbCrLf
            End If
        Next
    End With
    If strCopy <> "" Then
        VB.Clipboard.Clear
        VB.Clipboard.SetText strCopy
    End If
End Sub

Public Sub zlItemRef()
'功能：调用诊疗参考
    Dim lng诊疗项目ID As Long, i As Long

    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) <> 0 Then
            If .TextMatrix(.Row, COL_诊疗类别) = "E" And (RowIs配方行(.Row) Or RowIs检验行(.Row)) Then
                lng诊疗项目ID = Get诊疗项目ID(Val(.TextMatrix(.Row, COL_ID)), True)
            Else
                lng诊疗项目ID = Get诊疗项目ID(Val(.TextMatrix(.Row, COL_ID)), False)
            End If
        End If
    End With
    Call ShowClinicHelp(0, mfrmParent, lng诊疗项目ID)
End Sub

Public Sub zlPrintSetup()
    Call zlPrintSet
End Sub

Public Sub zlExcel()
    Call OutputList(3)
End Sub

Public Sub zlPreview()
    Call OutputList(2)
End Sub

Public Sub zlPrint()
    Call OutputList(1)
End Sub

Private Sub Form_Activate()
    picFocus.SetFocus '这样设置后本窗体内的焦点顺序才有效
    vsAdvice.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objMenu As Object
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu新开医嘱)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyM Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu修改医嘱)
    ElseIf KeyCode = vbKeyDelete Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu删除医嘱)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyT Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu皮试结果)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu发送医嘱)
    ElseIf KeyCode = vbKeyF2 Then '主界面定位病人
        Call mfrmParent.Form_KeyDown(vbKeyF2, 0): Exit Sub
    ElseIf KeyCode = vbKeyF3 Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu发送医嘱)
    ElseIf KeyCode = vbKeyF6 Then
        Call zlItemRef
    End If
    If Not objMenu Is Nothing Then
        If objMenu.Enabled And objMenu.Visible Then
            Call zlMenuClick(objMenu)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Call mfrmParent.Form_KeyPress(KeyAscii)
End Sub

Private Sub fraAdviceUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsAdvice.Height + y < 1000 Or vsAppend.Height - y < 60 Then Exit Sub
        fraAdviceUD.Top = fraAdviceUD.Top + y
        tabAppend.Top = tabAppend.Top + y
        vsAdvice.Height = vsAdvice.Height + y
        vsAppend.Top = vsAppend.Top + y
        vsAppend.Height = vsAppend.Height - y
        Me.Refresh
    End If
End Sub

Private Sub mfrmEdit_Unload(Cancel As Integer)
    If Not Cancel Then
        If frmOutAdviceEdit.mblnOK Then
            Call LoadAdvice
            Call ShowTotalMoney
        End If
        Set mfrmEdit = Nothing
        
        If mfrmParent.TabFile.SelectedItem.Key = "医嘱" Then
            Call BringWindowToTop(Me.Hwnd)
        End If
    End If
End Sub

Private Function CheckWindow() As Boolean
'功能：检查医嘱编辑窗口是否已经打开
    If Not mfrmEdit Is Nothing Then
        '当前窗口打开了
        MsgBox "医嘱编辑窗口已经打开，请先完成当前操作后再执行。", vbInformation, gstrSysName
        '定位到当前的窗口
        If mfrmEdit.WindowState = vbMinimized Then mfrmEdit.WindowState = vbNormal
        mfrmEdit.SetFocus
        Exit Function
    Else
        '其它窗口打开了
        If Not CheckAdviceWindow("门诊医嘱编辑") Then Exit Function
    End If
    CheckWindow = True
End Function

Private Sub FuncBillPrint(objMenu As Menu)
'功能：打印诊疗单据
    If objMenu.Tag = "" Then Exit Sub
    If ReportPrintSet(gcnOracle, glngSys, objMenu.Tag, mfrmParent) Then
        With vsAppend
            Call ReportOpen(gcnOracle, glngSys, objMenu.Tag, mfrmParent, "NO=" & .TextMatrix(.Row, cs单据号), "性质=" & Val(.TextMatrix(.Row, cs记录性质)), 2)
        End With
    End If
End Sub

Private Sub FuncAdviceSign()
'功能：对医嘱进行电子签名
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lng签名ID As Long, lng证书ID As Long
    Dim intRule As Integer
    
    If mint状态 <> 1 Then Exit Sub '在诊病人
    If gobjESign Is Nothing Then Exit Sub
    
    '获取签名医嘱源文
    intRule = ReadAdviceSignSource(1, mlng病人ID, mstr挂号单, strIDs, 0, mblnMoved, strSource, mlng前提ID)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "该病人目前没有可以签名的医嘱。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID)
    If strSign <> "" Then
        lng签名ID = zlDatabase.GetNextId("医嘱签名记录")
        strSQL = "zl_医嘱签名记录_Insert(" & lng签名ID & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & strIDs & "')"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        
        Call LoadAdvice '刷新界面
        MsgBox "已完成电子签名。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignErase()
'功能：取消医嘱的电子签名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If mint状态 <> 1 Then Exit Sub '在诊病人
    If gobjESign Is Nothing Then Exit Sub
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tabAppend.SelectedItem.Index <> 3 Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "当前选择的医嘱没有签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '作废签名不能取消
        If .Cell(flexcpData, .Row, 0) = 4 Then
            MsgBox "作废医嘱的签名不能取消。", vbInformation, gstrSysName
            Exit Sub
        End If
        '新开签名必须是在新开状态
        If .Cell(flexcpData, .Row, 0) = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) <> 1 Then
                MsgBox "由于医嘱已经发送或作废，该签名不能取消。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '不能取消医技下达的签名
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) <> 0 Then
            MsgBox "你不能取消医技科室下达医嘱的签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        '只能取消自已签的名
        If .TextMatrix(.Row, 2) <> UserInfo.姓名 Then
            MsgBox "该签名人不是你本人，不能取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("确实要取消这次签名吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        
        strSQL = "zl_医嘱签名记录_Delete(" & .RowData(.Row) & ")"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End With
    
    Call LoadAdvice '刷新界面
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignVerify()
'功能：校验医嘱的电子签名(可对已转移的数据)
    Dim strSource As String
    
    If gobjESign Is Nothing Then Exit Sub
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tabAppend.SelectedItem.Index <> 3 Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "当前选择的医嘱没有签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '获取签名医嘱源文
        If ReadAdviceSignSource(.Cell(flexcpData, .Row, 0), 0, 0, "", .RowData(.Row), mblnMoved, strSource) = 0 Then Exit Sub
        
        '验证签名
        Call gobjESign.VerifySignature(strSource, .RowData(.Row), 1)
    End With
End Sub

Private Sub FuncAdviceAdd()
'功能：新增医嘱
    If Not CheckWindow Then Exit Sub
    If mlng病人ID = 0 Then Exit Sub
    If mint状态 <> 1 Then Exit Sub '在诊病人
    
    Set mfrmEdit = frmOutAdviceEdit
    Call frmOutAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng病人ID, mstr挂号单, mlng前提ID)
End Sub

Private Sub FuncAdviceDel()
'删除：删除当前医嘱
'说明：在主界面删除,对检查组合,手术组合,中药配方,是整个删除,一并给药只删除当前药品
    Dim strSQL As String, lng医嘱ID As Long
    Dim blnGroup As Boolean, i As Long
    Dim lngRow As Long
    
    If mint状态 <> 1 Then Exit Sub '在诊病人
    
    With vsAdvice
        '检查是否可以删除
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        If lng医嘱ID = 0 Then
            MsgBox "该病人没有医嘱可以删除。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '医技下达的医嘱
        If Val(.TextMatrix(.Row, COL_前提ID)) <> mlng前提ID Then
            MsgBox "你不能删除该医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(.TextMatrix(.Row, COL_医嘱状态)) <> 1 Then
            MsgBox "当前选择的医嘱已经发送或作废，不能删除。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已签名的医嘱不能删除
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            MsgBox "当前选择的医嘱已经签名，不能删除。请先取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If InStr(",5,6,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Then
            If .Row - 1 >= .FixedRows Then
                If Val(.TextMatrix(.Row - 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then blnGroup = True
            End If
            If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                If Val(.TextMatrix(.Row + 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then blnGroup = True
            End If
            If blnGroup Then
                If MsgBox("医嘱""" & .TextMatrix(.Row, COL_医嘱内容) & """与其它药品一并给药,确实要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        If Not blnGroup Then
            If MsgBox("确实要删除医嘱""" & .TextMatrix(.Row, COL_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        strSQL = "ZL_病人医嘱记录_Delete(" & lng医嘱ID & ",1)"
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    With vsAdvice
        '界面上直接删除
        .Redraw = False
        
        '删除一并给药第一行时的显示处理
        If blnGroup And .Row + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(.Row, COL_相关ID)) = Val(.TextMatrix(.Row + 1, COL_相关ID)) Then
                If .TextMatrix(.Row, COL_开始时间) <> "" And .TextMatrix(.Row + 1, COL_开始时间) = "" Then
                    .TextMatrix(.Row + 1, COL_开始时间) = .TextMatrix(.Row, COL_开始时间)
                    .TextMatrix(.Row + 1, COL_频率) = .TextMatrix(.Row, COL_频率)
                    .TextMatrix(.Row + 1, COL_用法) = .TextMatrix(.Row, COL_用法)
                End If
            End If
        End If
        
        lngRow = .Row
        .RemoveItem .Row
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If

        Call .ShowCell(.Row, .Col)
        .Redraw = True
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col) '颜色及附表更新
    End With
    Call ShowTotalMoney
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceRevoke()
'删除：当前医嘱作废(一组医嘱作废)
    Dim strSQL As String, lng医嘱ID As Long
    Dim lng证书ID As Long, lng签名ID As Long
    Dim strSign As String, intRule As Integer
    Dim strSource As String, strIDs As String
    
    If mint状态 <> 1 Then Exit Sub '在诊病人
    
    With vsAdvice
        '检查是否可以作废
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        If lng医嘱ID = 0 Then
            MsgBox "该病人没有医嘱可以作废。", vbInformation, gstrSysName
            Exit Sub
        End If
                
        If Val(.TextMatrix(.Row, COL_前提ID)) <> mlng前提ID Then
            MsgBox "你不能作废该医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(.TextMatrix(.Row, COL_医嘱状态)) <> 8 Then
            MsgBox "当前选择的医嘱尚未发送或已经作废。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已有费用转出不允许作废
        If MovedByDate(.Cell(flexcpData, .Row, COL_发送时间)) Then
            If MovedBySend(lng医嘱ID) Then
                MsgBox "该医嘱的费用已经全部或部份转出到后备数据库，不允许操作。" & vbCrLf & _
                    "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '电子签名检查和提示
        If Val(.TextMatrix(.Row, COL_签名否)) = "1" Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "作废已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能作废。", vbInformation, gstrSysName
                Else
                    MsgBox "作废已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能作废。", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
            strSign = vbCrLf & vbCrLf & "提示：该医嘱已经签名，作废时你需要再次签名。"
        End If
                
        '检查作废医嘱对应的费用结帐情况
        If Not CheckAdviceBalanceRevoke(lng医嘱ID) Then Exit Sub
                
        If RowIn一并给药(.Row, 0, 0) Then
            If MsgBox("该组一并给药的医嘱将会一起作废，确实要作废吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("确实要作废医嘱""" & .TextMatrix(.Row, COL_医嘱内容) & """吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        strSQL = "ZL_病人医嘱记录_作废(" & lng医嘱ID & ")"
        
        '作废时进行电子签名
        If strSign <> "" Then
            '获取签名医嘱源文
            strIDs = lng医嘱ID
            intRule = ReadAdviceSignSource(4, mlng病人ID, mstr挂号单, strIDs, 0, mblnMoved, strSource)
            If intRule = 0 Then Exit Sub
            If strSource = "" Then
                MsgBox "不能读取需要作废的已签名医嘱源文内容。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID)
            If strSign <> "" Then
                lng签名ID = zlDatabase.GetNextId("医嘱签名记录")
                strSign = "zl_医嘱签名记录_Insert(" & lng签名ID & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & strIDs & "')"
            Else
                Exit Sub
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    If strSign <> "" Then
        zlDatabase.ExecuteProcedure strSign, Me.Name
    End If
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    Call LoadAdvice '刷新界面
    Call ShowTotalMoney
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceModi()
'功能：修改当前医嘱
    Dim lng医嘱ID As Long
    
    If Not CheckWindow Then Exit Sub
    
    If mlng病人ID = 0 Then Exit Sub
    If mint状态 <> 1 Then Exit Sub '在诊病人
    
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        If lng医嘱ID = 0 Then Exit Sub
        
        '医技下达的医嘱
        If Val(.TextMatrix(.Row, COL_前提ID)) <> mlng前提ID Then
            MsgBox "你不能修改该医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已校对或已废止
        If Val(.TextMatrix(.Row, COL_医嘱状态)) <> 1 Then
            MsgBox "当前选择的医嘱已经发送或作废，不能修改。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已签名的医嘱不能修改
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            MsgBox "当前选择的医嘱已经签名，不能修改。请先取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Set mfrmEdit = frmOutAdviceEdit
        Call frmOutAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng病人ID, mstr挂号单, mlng前提ID, _
            Val(.TextMatrix(.Row, COL_婴儿ID)), lng医嘱ID)
    End With
End Sub

Private Sub FuncAdviceTest()
'功能：填写皮试结果
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim v结果 As VbMsgBoxResult, str结果 As String
    
    If mlng病人ID = 0 Then Exit Sub
    If mint状态 <> 1 Then Exit Sub '在诊病人
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Then Exit Sub
    If Not (vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "E" And vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型) = "1") Then
        MsgBox "当前医嘱内容不是过敏试验项目。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) <> 0 Then
        MsgBox "你不能给该过敏试验填写结果。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 4 Then
        MsgBox "该过敏试验医嘱已经作废，不能填写结果。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 1 Then
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_皮试) = "免试" Then
            If MsgBox("该过敏试验医嘱已经标记为免试，要清除免试标记吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            str结果 = ""
        Else
            If MsgBox("该过敏试验医嘱尚未发送，不允许填写过敏试验结果。" & vbCrLf & vbCrLf & _
                "但可以标记为免试，同时该医嘱将不会发送。要标记为免试吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            str结果 = "免试"
        End If
    Else
        '检查对应的医嘱是否已经发送
        If mbln自动皮试 Then
            If AdviceSended(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))) Then
                MsgBox "该皮试对应的药品已经发送，不能再更改皮试结果。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_皮试) <> "" Then
            If MsgBox("该过敏试验医嘱已经填写了结果，要重新填写吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        v结果 = frmMsgBox.ShowMsgBox(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱内容) & "：^^请根据过敏试验结果选择相应的按钮操作。", mfrmParent, , 1)
        If v结果 = vbCancel Then Exit Sub
        str结果 = IIF(v结果 = vbYes, "(+)", "(-)")
    End If
    
    strSQL = "ZL_病人医嘱记录_皮试(" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & ",'" & str结果 & "')"
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    vsAdvice.TextMatrix(vsAdvice.Row, COL_皮试) = str结果
    If str结果 = "(+)" Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_皮试) = vbRed
    ElseIf str结果 = "(-)" Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_皮试) = vbBlue
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSended(ByVal lng医嘱ID As Long) As Boolean
'功能：判断皮试对应的医嘱是否已经发送
'参数：lng医嘱ID=皮试医嘱的ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '已作废的不管
    strSQL = "Select 诊疗项目ID From 病人医嘱记录 Where ID=[3]"
    strSQL = "Select A.ID From 病人医嘱记录 A,诊疗用法用量 B" & _
        " Where Rownum<2 And A.诊疗类别 IN('5','6') And A.医嘱状态=8" & _
        " And A.诊疗项目ID=B.项目ID And B.性质=0 And B.用法ID=(" & strSQL & ")" & _
        " And A.病人ID+0=[1] And A.挂号单=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mstr挂号单, lng医嘱ID)
    AdviceSended = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceSend()
'功能：发送病人医嘱(可以设置计价项目)

    If mlng病人ID = 0 Then Exit Sub
    If mint状态 <> 1 Then Exit Sub '在诊病人
    
    If frmOutAdviceSend.ShowMe(Me, mstrPrivs, mlng病人ID, mstr挂号单, mlng前提ID) Then
        Call LoadAdvice
        Call ShowTotalMoney
    End If
End Sub

Private Sub tabAppend_Click()
    If Val(vsAppend.Tag) = tabAppend.SelectedItem.Index Then Exit Sub
    
    If Visible Then
        Call SaveFlexState(vsAppend, App.ProductName & "\" & Me.Name)
    End If
        
    vsAppend.Tag = tabAppend.SelectedItem.Index
    If tabAppend.SelectedItem.Index = 1 Then
        Call InitPriceTable
    ElseIf tabAppend.SelectedItem.Index = 2 Then
        Call InitSendTable
    ElseIf tabAppend.SelectedItem.Index = 3 Then
        Call InitSignTable
    End If
    
    If Visible Then
        Call RestoreFlexState(vsAppend, App.ProductName & "\" & Me.Name)
    End If
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    If Visible Then vsAdvice.SetFocus
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    
    If NewRow = OldRow Then Exit Sub
    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, COL_开始时间)
    End If
    If vsAdvice.Redraw <> flexRDNone Then
        If Val(vsAdvice.TextMatrix(NewRow, COL_ID)) <> 0 Then
            '显示医嘱附加表格的内容
            If mfrmParent.mnuViewAdviceAppend.Checked Then
                If tabAppend.SelectedItem.Index = 1 Then
                    Call ShowPrice(NewRow)
                ElseIf tabAppend.SelectedItem.Index = 2 Then
                    Call ShowSendList(NewRow)
                ElseIf tabAppend.SelectedItem.Index = 3 Then
                    Call ShowSignList(NewRow)
                End If
            End If
        ElseIf mfrmParent.mnuViewAdviceAppend.Checked Then
            Call ClearAppendData
            vsAppend.Row = vsAppend.FixedRows
        End If
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = COL_医嘱内容 Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_皮试 Then
            Cancel = True
        ElseIf Col = COL_警示 Then 'Pass
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '擦除固定列中的表格线
            SetBkColor hDC, SysColor2RGB(.BackColorFixed)

            '仅左边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅上边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅下边表格线
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅右边表格线
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            '擦除一并给药相关行列的边线及内容
            lngLeft = COL_开始时间: lngRight = COL_开始时间
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_频率: lngRight = COL_用法
                If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            End If
            
            If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
            
            vRect.Left = Left '擦除左边表格线
            vRect.Right = Right - 1 '保留右边表格线
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '底行保留下边线
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
                '为了支持预览输出
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    
    If Button = 2 And mfrmParent.mnuAdvice.Visible Then PopupMenu mfrmParent.mnuAdvice, 2
End Sub

Private Function GetPatiInfo() As String
'功能：读取病人信息串(用于打印)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '执行部门(号别科室)即病人科室
    strSQL = "Select B.姓名,B.性别,B.年龄,B.门诊号," & _
        " B.险类,B.就诊诊室,C.名称 as 执行部门,A.执行部门ID,A.登记时间" & _
        " From 病人挂号记录 A,病人信息 B,部门表 C" & _
        " Where A.NO=[2] And A.病人ID+0=[1]" & _
        " And A.病人ID=B.病人ID And A.执行部门ID=C.ID"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mstr挂号单)
    
    GetPatiInfo = _
        "姓名：" & rsTmp!姓名 & " 性别：" & Nvl(rsTmp!性别) & _
        " 年龄：" & Nvl(rsTmp!年龄) & " 门诊号：" & Nvl(rsTmp!门诊号) & _
        " 挂号：" & Format(rsTmp!登记时间, "MM-dd HH:mm") & _
        " 科室：" & rsTmp!执行部门 & " 诊室：" & Nvl(rsTmp!就诊诊室)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String
    
    If mlng病人ID = 0 Then Exit Sub
    
    '表头
    objOut.Title.Text = "病人医嘱清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    Set objRow = New zlTabAppRow
    objRow.Add GetPatiInfo
    objOut.UnderAppRows.Add objRow
    
    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = vsAdvice
    
    '输出
    vsAdvice.Redraw = False
    lngRow = vsAdvice.Row: lngCol = vsAdvice.Col
    
    strWidth = ""
    For i = 0 To vsAdvice.FixedCols - 1
        strWidth = strWidth & "," & vsAdvice.ColWidth(i)
        vsAdvice.ColWidth(i) = 0
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsAdvice.FixedCols - 1
        vsAdvice.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    
    vsAdvice.Row = lngRow: vsAdvice.Col = lngCol
    vsAdvice.Redraw = True
End Sub

Private Sub Form_Load()
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    
    '电子签名记录
    If gobjESign Is Nothing Then
        tabAppend.Tabs.Remove 3
    End If
    
    Call InitAdviceTable
    Call tabAppend_Click
    Call RestoreWinState(Me, App.ProductName)
    
    '自动处理皮试
    mbln自动皮试 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "自动处理皮试", 0)) <> 0

    fraAdviceUD.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    tabAppend.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    vsAppend.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    
    Set mfrmEdit = Nothing
    
    Call InitSysPar '初始化系统参数
End Sub

Private Sub Form_Resize()
    Dim PriceH As Long

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    PriceH = IIF(vsAppend.Visible, vsAppend.Height + fraAdviceUD.Height + tabAppend.Height, 0)
    
    vsAdvice.Left = 0
    vsAdvice.Top = 0
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - PriceH
    
    fraAdviceUD.Left = 0
    fraAdviceUD.Top = vsAdvice.Top + vsAdvice.Height
    fraAdviceUD.Width = Me.ScaleWidth
    
    tabAppend.Left = 0
    tabAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    tabAppend.Width = Me.ScaleWidth
    
    vsAppend.Left = 0
    vsAppend.Top = tabAppend.Top + tabAppend.Height
    vsAppend.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmEdit = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub ClearAppendData()
'功能：清除附加表格数据
    vsAppend.Rows = vsAppend.FixedRows
    vsAppend.Rows = vsAppend.FixedRows + 1
End Sub

Private Sub InitPriceTable()
'功能：初始化计价清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "计价医嘱,2000,1;类别,650,1;收费项目,2500,1;单位,500,4;数量,500,1;单价,800,7;执行科室,1000,1;费用类型,800,1;从项,450,4"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCol(0) = False
        .MergeCol(1) = False
    End With
End Sub

Private Sub InitSendTable()
'功能：初始化发送清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "发送号;发送时间,1080,1;发送医嘱,1800,1;单据号,850,1;收费项目,1800,1;发送数次,850,1;计费状态,850,1;执行状态,850,1;执行科室,850,1;发送人,800,1;记录性质"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCol(0) = True
        .MergeCol(1) = True
    End With
End Sub

Private Sub InitSignTable()
'功能：初始化计价清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "签名类型,1150,1;签名时间,1900,1;签名人,800,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCol(0) = False
        .MergeCol(1) = False
    End With
End Sub

Private Sub ClearAdviceData()
'功能：清除医嘱清单数据
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Editable = flexEDNone
End Sub

Private Sub InitAdviceTable()
'功能：初始化医嘱清单格式
    Dim arrHead As Variant, strHead As String, i As Long

    strHead = "ID;相关ID;组ID;组号;婴儿ID;医嘱状态;诊疗类别;操作类型;毒理分类;标志;" & _
        ",240,4;开始时间,1080,1;医嘱内容,3000,1;,375,4;总量,850,1;单量,850,1;频率,1000,1;" & _
        "用法,1000,1;医生嘱托,1000,1;执行时间,1000,1;执行科室,850,1;执行性质,850,1;" & _
        "开嘱医生,850,1;开嘱时间,1080,1;发送人,850,1;发送时间,1080,1;" & _
        "单据ID;申请项;报告项;申请ID;前提ID;签名否"
    arrHead = Split(strHead, ";")
    With vsAdvice
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
        .ColHidden(COL_警示) = Not (gblnPass And InStr(mstrPrivs, "合理用药监测") > 0) 'Pass
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .ColWidth(0) = 9 * Screen.TwipsPerPixelX
        .ColWidth(1) = 11 * Screen.TwipsPerPixelX
    End With
End Sub

Private Function LoadAdvice() As Boolean
'功能：根据当前界面设置读取并显示医嘱清单
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngTop As Long
    Dim strFormat As String, strTmp As String
    Dim bln给药途径 As Boolean, bln中药用法 As Boolean
        Dim bln采集方法 As Boolean, bln申请项 As Boolean, bln已申请 As Boolean
    Dim blnFirst As Boolean, lng医嘱ID As Long
    Dim strBill As String, i As Long, j As Long
    
    If mlng病人ID = 0 Then Exit Function
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) '记录当前行
        
    '诊疗单据：对应诊疗单据,及申请项,报告项
    strBill = "Select A.ID as 医嘱ID,B.病历文件ID as 单据ID," & _
        " Max(Decode(C.填写时机,1,1,0)) as 申请项," & _
        " Max(Decode(C.填写时机,2,1,0)) as 报告项" & _
        " From 病人医嘱记录 A,诊疗单据应用 B,病历文件组成 C" & _
        " Where A.诊疗项目ID=B.诊疗项目ID And B.应用场合=1 And B.病历文件ID=C.病历文件ID(+)" & _
        " And A.病人ID+0=[1] And A.挂号单=[2]" & _
        " And Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL)" & _
        " Group by A.ID,B.病历文件ID"
        
    '医嘱记录：不含附加手术,手术麻醉,检查部位,中药煎法
    strSQL = _
        "Select /*+ RULE */ A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,Nvl(X.序号,A.序号) as 组号," & _
            " Nvl(A.婴儿,0) as 婴儿ID,A.医嘱状态,A.诊疗类别,B.操作类型,C.毒理分类,A.紧急标志 as 标志," & _
            " A.审查结果,To_Char(A.开始执行时间,'MM-DD HH24:MI') as 开始时间,A.医嘱内容,A.皮试结果 as 皮试," & _
            " Decode(A.总给予量,NULL,NULL,Decode(A.诊疗类别,'E',Decode(B.操作类型,'4',A.总给予量||'付',A.总给予量||B.计算单位),'5',Round(A.总给予量/D.门诊包装,5)||D.门诊单位,'6',Round(A.总给予量/D.门诊包装,5)||D.门诊单位,A.总给予量||B.计算单位)) as 总量," & _
            " Decode(A.单次用量,NULL,NULL,A.单次用量||B.计算单位) as 单量," & _
            " A.执行频次 as 频率,Decode(A.诊疗类别,'E',Decode(Instr('246',Nvl(B.操作类型,'0')),0,NULL,B.名称),NULL) as 用法," & _
            " A.医生嘱托,A.执行时间方案 as 执行时间,Nvl(E.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'<院外执行>')) as 执行科室," & _
            " Decode(Instr('567E',A.诊疗类别),0,NULL,A.执行性质) as 执行性质," & _
            " A.开嘱医生,To_Char(A.开嘱时间,'MM-DD HH24:MI') as 开嘱时间," & _
            " A.停嘱医生 as 发送人,A.停嘱时间 as 发送时间," & _
            " Y.单据ID,Y.申请项,Y.报告项,A.申请ID,A.前提ID," & _
            " Decode(S.签名ID,NULL,0,1) as 签名否" & _
        " From 病人医嘱记录 A,部门表 E,药品特性 C,药品规格 D,诊疗项目目录 B,病人医嘱状态 S,病人医嘱记录 X,(" & strBill & ") Y" & _
        " Where A.诊疗项目ID=B.ID And A.执行科室ID=E.ID(+) And A.诊疗项目ID=C.药名ID(+)" & _
            " And Nvl(A.医嘱期效,0)=1 And A.收费细目ID=D.药品ID(+) And A.相关ID=X.ID(+)" & _
            " And Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL)" & _
            " And A.开始执行时间 is Not NULL And A.病人来源<>3 And A.ID=S.医嘱ID And S.操作类型=1" & _
            " And A.ID=Y.医嘱ID(+) And A.病人ID+0=[1] And A.挂号单=[2]" & _
            IIF(mlng前提ID = 0 Or mblnShowAll, "", " And A.前提ID=[3]") & _
        " Order by Nvl(A.婴儿,0),组号,组ID,A.序号"
        
    If mblnMoved Then '挂号单与医嘱同个数据库
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mstr挂号单, mlng前提ID)
    
    If Not rsTmp.EOF Then
        With vsAdvice
            .Redraw = False
            
            '绑定时按设计时的FormatString恢复一些缺省值(固定行列数，固定行列文字及行列对齐,尺寸,可见)
            'FormatString在运行时赋值无效
            '如果AutoResize=True,则所有列宽或行高被自动调整(根据AutoSizeMode)
            '如果WordWrap=True,则行高会被自动调整
            .WordWrap = False
            strFormat = GetColFormat(vsAdvice)
            Call ClearAdviceData
            .ScrollBars = flexScrollBarNone
            Set .DataSource = rsTmp
            .ScrollBars = flexScrollBarBoth
            If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
                gcnOracle.Errors.Clear '怪,绑定时固定有此错误
            End If
            Call SetColFormat(vsAdvice, strFormat)
            .TextMatrix(0, COL_皮试) = ""
            .TextMatrix(0, COL_警示) = "" 'Pass
            
            '自动调整行高
            .WordWrap = True
            .AutoSize COL_医嘱内容
            
            '处理每行医嘱
            i = .FixedRows
            Do While i <= .Rows - 1
                '处理发送时间
                If .TextMatrix(i, COL_发送时间) <> "" Then
                    .Cell(flexcpData, i, COL_发送时间) = .TextMatrix(i, COL_发送时间)
                    .TextMatrix(i, COL_发送时间) = Format(.TextMatrix(i, COL_发送时间), "MM-dd HH:mm")
                End If
                
                '成药及中药的一些处理
                bln给药途径 = False: bln中药用法 = False
                bln采集方法 = False: bln申请项 = False: bln已申请 = False '仅用于检验组合
                If .TextMatrix(i, COL_诊疗类别) = "E" Then
                    If Val(.TextMatrix(i - 1, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                            bln给药途径 = True
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    '显示成药的给药途径
                                    .TextMatrix(j, COL_用法) = .TextMatrix(i, COL_用法)
                                    '显示成药的执行性质
                                    If Val(.TextMatrix(j, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                        .TextMatrix(j, COL_执行性质) = "自备药"
                                    ElseIf Val(.TextMatrix(j, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                        .TextMatrix(j, COL_执行性质) = "离院带药"
                                    Else
                                        .TextMatrix(j, COL_执行性质) = ""
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                            bln中药用法 = .TextMatrix(i - 1, COL_诊疗类别) = "7" '中药用法行
                            bln采集方法 = .TextMatrix(i - 1, COL_诊疗类别) = "C" '采集方法行

                            '显示中药配方或检验组合的执行科室
                            .TextMatrix(i, COL_执行科室) = .TextMatrix(i - 1, COL_执行科室)
                            
                            If bln中药用法 Then
                                '显示中药配方执行性质
                                If Val(.TextMatrix(i - 1, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                    .TextMatrix(i, COL_执行性质) = "自备药"
                                ElseIf Val(.TextMatrix(i - 1, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                    .TextMatrix(i, COL_执行性质) = "离院带药"
                                Else
                                    .TextMatrix(i, COL_执行性质) = ""
                                End If
                            Else
                                .TextMatrix(i, COL_执行性质) = ""
                            End If
                            
                            '删除单味中药行,以及检验组合中的检验项目;同时判断检验申请
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    If .TextMatrix(j, COL_诊疗类别) = "C" Then
                                        If Val(.TextMatrix(j, COL_申请项)) = 1 Then
                                            bln申请项 = True
                                            If Val(.TextMatrix(j, COL_申请ID)) <> 0 Then
                                                bln已申请 = True
                                            End If
                                        End If
                                    End If
                                    .RemoveItem j: i = i - 1
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        .TextMatrix(i, COL_执行性质) = ""
                    End If
                End If
                                                                
                '处理可见行的的一些标识:排开不可见但暂时未删除的行
                If Not bln给药途径 And .TextMatrix(i, COL_诊疗类别) <> "7" Then
                    
                    '行高：为了支持zl9PrintMode:Resize之后,取RowHeight可能小于RowHeightMin
                    If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                    
                    '处理小数点问题,暂未想到办法
                    If Left(.TextMatrix(i, COL_总量), 1) = "." Then
                        .TextMatrix(i, COL_总量) = "0" & .TextMatrix(i, COL_总量)
                    End If
                    If Left(.TextMatrix(i, COL_单量), 1) = "." Then
                        .TextMatrix(i, COL_单量) = "0" & .TextMatrix(i, COL_单量)
                    End If
                    
                    '可申请医嘱标识(不管药品及相关医嘱,且只管主要医嘱)
                    If Not bln中药用法 And InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) = 0 Then
                        If bln采集方法 Then '利用前面取的结果
                            If bln申请项 Then
                                If Not bln已申请 Then
                                    Set .Cell(flexcpPicture, i, COL_F申请) = imgFlag.ListImages("未申请").Picture
                                Else
                                    Set .Cell(flexcpPicture, i, COL_F申请) = imgFlag.ListImages("已申请").Picture
                                End If
                            End If
                        ElseIf Val(.TextMatrix(i, COL_申请项)) = 1 Then
                            If Val(.TextMatrix(i, COL_申请ID)) = 0 Then
                                Set .Cell(flexcpPicture, i, COL_F申请) = imgFlag.ListImages("未申请").Picture
                            Else
                                Set .Cell(flexcpPicture, i, COL_F申请) = imgFlag.ListImages("已申请").Picture
                            End If
                        End If
                    End If
                    
                    '医嘱颜色
                    If Val(.TextMatrix(i, COL_医嘱状态)) = 4 Then
                        '已作废(发送后作废)
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '灰色
                    ElseIf Val(.TextMatrix(i, COL_医嘱状态)) = 8 Then
                        '已发送(发送后自动停止)
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000 '深蓝
                        If lngTop = 0 Then lngTop = i
                    Else
                        If lngTop = 0 Then lngTop = i
                    End If
                    
                    '毒麻精药品标识:中药配方及组成味中药不处理
                    If .TextMatrix(i, COL_毒理分类) <> "" Then
                        If InStr(",麻醉药,毒性药,精神药,", .TextMatrix(i, COL_毒理分类)) > 0 Then
                            .Cell(flexcpFontBold, i, COL_医嘱内容) = True
                        End If
                    End If
                    
                    '皮试结果标识
                    If .TextMatrix(i, COL_皮试) = "(+)" Then
                        .Cell(flexcpForeColor, i, COL_皮试) = vbRed
                    ElseIf .TextMatrix(i, COL_皮试) = "(-)" Then
                        .Cell(flexcpForeColor, i, COL_皮试) = vbBlue
                    End If
                    
                    '紧急标志:一并给药只显示在第一行
                    blnFirst = True
                    If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            blnFirst = False
                        End If
                    End If
                    If blnFirst Then
                        If Val(.TextMatrix(i, COL_标志)) = 1 Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("紧急").Picture
                        ElseIf Val(.TextMatrix(i, COL_标志)) = 2 Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("补录").Picture
                        End If
                    End If
                    
                    'Pass:根据审查结果显示警示灯
                    If .TextMatrix(i, COL_警示) <> "" Then
                        Set .Cell(flexcpPicture, i, COL_警示) = imgPass.ListImages(Val(.TextMatrix(i, COL_警示)) + 1).Picture
                        .TextMatrix(i, COL_警示) = ""
                    End If
                    
                    '电子签名标识
                    If Val(.TextMatrix(i, COL_签名否)) = 1 Then
                        Set .Cell(flexcpPicture, i, COL_医嘱内容) = imgSign.ListImages(1).Picture
                    End If
                End If
                
                If bln给药途径 Then
                    .RemoveItem i
                Else
                    i = i + 1
                End If
            Loop
            
            '固定列图标对齐:设置为中对齐,不然擦边框时可能有问题
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '电子签名图标对齐
            .Cell(flexcpPictureAlignment, .FixedRows, COL_医嘱内容, .Rows - 1, COL_医嘱内容) = 0
            .Redraw = True
        End With
    Else
        Call ClearAdviceData
        Call ClearAppendData
    End If
        
    '缺省定位
    If lng医嘱ID <> 0 Then
        lng医嘱ID = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
        If lng医嘱ID <> -1 Then vsAdvice.Row = lng医嘱ID
    End If
    If lng医嘱ID = -1 Or lng医嘱ID = 0 Then
        If lngTop <> 0 Then
            vsAdvice.Row = lngTop
            vsAdvice.TopRow = lngTop
        Else
            vsAdvice.Row = vsAdvice.FixedRows
        End If
    End If
    vsAdvice.Col = vsAdvice.FixedCols
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    vsAdvice.Refresh
    Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = True
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否中药配方行
'说明：指定行为显示行,且类别="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From 病人医嘱记录 Where Rownum=1 And 诊疗类别='7' And 相关ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs配方行 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否检验组合行
'说明：指定行为显示行,且类别="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From 病人医嘱记录 Where Rownum=1 And 诊疗类别='C' And 相关ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs检验行 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowPrice(ByVal lngRow As Long) As Boolean
'功能：读取指定医嘱的计价,并根据当前的诊疗收费关系进行更新
    Dim rs诊疗项目 As New ADODB.Recordset
    Dim rs收费细目 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim str医嘱IDs As String, str收费细目IDs As String
    Dim strSQL As String, i As Long, j As Long
    Dim bln配方行 As Boolean, bln检验行 As Boolean, blnLoad As Boolean
    Dim lng病人科室ID As Long, lng执行科室ID As Long
    Dim dblPrice As Double
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .MergeCells = flexMergeNever
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowPrice = True: Exit Function
        End If
        If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            bln配方行 = RowIs配方行(lngRow)
            bln检验行 = RowIs检验行(lngRow)
        End If
                                    
        blnLoad = True
        
        '药品的计价
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '中,西成药:可能按规格下医嘱,计算1个门诊包装的单价
            strSQL = "Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,NULL as 标本部位,C.ID as 收费细目ID," & _
                " B.门诊包装,B.门诊单位 as 计算单位,1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*B.门诊包装 as 单价," & _
                " A.执行科室ID,0 as 从项" & _
                " From 病人医嘱记录 A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.诊疗项目ID=B.药名ID And B.药品ID=C.ID And Nvl(A.执行性质,0)<>5" & _
                " And (A.收费细目ID is NULL Or A.收费细目ID=B.药品ID)" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.服务对象 IN(1,3) And D.收费细目ID=C.ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
                
                '仅一并给药(如果是)的第一成药行才显示给药途径的计价
                blnLoad = Val(vsAdvice.TextMatrix(lngRow - 1, COL_相关ID)) <> Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
        ElseIf bln配方行 Then
            '中草药:一定对应有规格记录且填写了收费细目ID
            strSQL = "Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,NULL as 标本部位,C.ID as 收费细目ID," & _
                " B.门诊包装,B.门诊单位 as 计算单位,1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*B.门诊包装 as 单价," & _
                " A.执行科室ID,0 as 从项" & _
                " From 病人医嘱记录 A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where A.诊疗类别='7' And A.相关ID=[1]" & _
                " And A.收费细目ID=B.药品ID And A.收费细目ID=C.ID And C.服务对象 IN(1,3)" & _
                " And D.收费细目ID=C.ID And Nvl(A.执行性质,0)<>5" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
        End If
        
        '读取现有计价(取最新价格)：除药品外的计价,包含相关医嘱计价
        '不计价,手工计价的医嘱不读取
        '用Union方式可以利用索引
        If blnLoad Then
            '不是新开的医嘱，根据病人医嘱计价提取
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位," & _
                " B.收费细目ID,1 as 门诊包装,C.计算单位,B.数量,Decode(C.是否变价,1,B.单价,Sum(D.现价)) as 单价," & _
                " Nvl(B.执行科室ID,A.执行科室ID) as 执行科室ID,Nvl(B.从项,0) as 从项" & _
                " From 病人医嘱记录 A,病人医嘱计价 B,收费项目目录 C,收费价目 D" & _
                " Where A.诊疗类别 Not IN('5','6','7') And A.ID=B.医嘱ID" & _
                " And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5) And B.收费细目ID=C.ID And B.收费细目ID=D.收费细目ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                " And (A.ID=[1] Or A.ID=[2] Or A.相关ID=[1])" & _
                " Group by A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,B.收费细目ID," & _
                " C.计算单位,B.数量,C.是否变价,B.单价,Nvl(B.执行科室ID,A.执行科室ID),Nvl(B.从项,0)"
            '新开的医嘱，根据诊疗收费关系提取(非药变价显示为0)
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,B.收费项目ID," & _
                " 1 as 门诊包装,C.计算单位,B.收费数量 as 数量,Decode(C.是否变价,1,0,Sum(D.现价)) as 单价," & _
                " A.执行科室ID,Nvl(B.从属项目,0) as 从项" & _
                " From 病人医嘱记录 A,诊疗收费关系 B,收费项目目录 C,收费价目 D" & _
                " Where A.诊疗类别 Not IN('5','6','7') And A.医嘱状态 IN(1,2) And A.诊疗项目ID=B.诊疗项目ID" & _
                " And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5) And B.收费项目ID=C.ID And B.收费项目ID=D.收费细目ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) And C.服务对象 IN(1,3)" & _
                " And (A.ID=[1] Or A.ID=[2] Or A.相关ID=[1])" & _
                " Group by A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,B.收费项目ID," & _
                " C.计算单位,B.收费数量,C.是否变价,A.执行科室ID,Nvl(B.从属项目,0)"
        End If
        strSQL = strSQL & " Order by 序号,从项"
        
        If mblnMoved Then '挂号单与医嘱在同个数据库
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱计价", "H病人医嘱计价")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)))
        
        '显示计价内容
        If Not rsTmp.EOF Then
            '确定显示行数
            .Rows = .FixedRows + rsTmp.RecordCount
            
            '获取诊疗项目,收费细目信息
            For i = 1 To rsTmp.RecordCount
                str医嘱IDs = str医嘱IDs & "," & rsTmp!ID
                str收费细目IDs = str收费细目IDs & " Union ALL Select " & rsTmp!收费细目ID & " From Dual"
                rsTmp.MoveNext
            Next
            str医嘱IDs = Mid(str医嘱IDs, 2)
            str收费细目IDs = Mid(str收费细目IDs, 12)
                        
            strSQL = "Select B.ID,B.类别,C.名称 as 类别名称,B.名称,B.标本部位" & _
                " From 病人医嘱记录 A,诊疗项目目录 B,诊疗项目类别 C" & _
                " Where A.ID IN(" & str医嘱IDs & ") And A.诊疗项目ID=B.ID And B.类别=C.编码"
                
            If mblnMoved Then '挂号单与医嘱在同个数据库
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            End If
            Call zlDatabase.OpenRecordset(rs诊疗项目, strSQL, Me.Name) 'In
            
            strSQL = "Select A.ID,A.类别,B.名称 as 类别名称,A.编码," & _
                " A.名称,A.规格,A.产地,A.费用类型,A.是否变价" & _
                " From 收费项目目录 A,收费项目类别 B" & _
                " Where A.类别=B.编码 And A.ID IN(" & str收费细目IDs & ")"
            strSQL = "Select A.ID,A.类别,A.类别名称,A.编码,Nvl(B.名称,A.名称) as 名称," & _
                " A.规格,A.产地,A.费用类型,A.是否变价,C.跟踪在用" & _
                " From (" & strSQL & ") A,收费项目别名 B,材料特性 C" & _
                " Where A.ID=C.材料ID(+) And A.ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIF(gbln商品名, 3, 1)
            Call zlDatabase.OpenRecordset(rs收费细目, strSQL, Me.Name) 'In
            
            '显示每行内容
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                rs诊疗项目.Filter = "ID=" & rsTmp!诊疗项目ID
                rs收费细目.Filter = "ID=" & rsTmp!收费细目ID
                
                '计价医嘱
                If InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                    .TextMatrix(i, 0) = "药品医嘱-" & rs诊疗项目!名称
                ElseIf rsTmp!诊疗类别 = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                    .TextMatrix(i, 0) = "给药途径-" & rs诊疗项目!名称
                ElseIf rsTmp!诊疗类别 = "E" And (bln配方行 Or bln检验行) Then
                    If bln检验行 Then
                        .TextMatrix(i, 0) = "采集方法-" & rs诊疗项目!名称
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        .TextMatrix(i, 0) = "中药煎法-" & rs诊疗项目!名称
                    Else
                        .TextMatrix(i, 0) = "中药用法-" & rs诊疗项目!名称
                    End If
                ElseIf Not IsNull(rsTmp!相关ID) Then
                    If rsTmp!诊疗类别 = "C" Then
                        .TextMatrix(i, 0) = "检验项目-" & rs诊疗项目!名称
                    ElseIf rsTmp!诊疗类别 = "D" Then
                        .TextMatrix(i, 0) = "检查部位-" & Nvl(rsTmp!标本部位)
                    ElseIf rsTmp!诊疗类别 = "F" Then
                        .TextMatrix(i, 0) = "附加手术-" & rs诊疗项目!名称
                    ElseIf rsTmp!诊疗类别 = "G" Then
                        .TextMatrix(i, 0) = "麻醉项目-" & rs诊疗项目!名称
                    End If
                Else
                    .TextMatrix(i, 0) = rs诊疗项目!类别名称 & "医嘱-" & rs诊疗项目!名称
                End If
                
                '类别
                .TextMatrix(i, 1) = rs收费细目!类别名称
                '收费项目:规格/产地
                .TextMatrix(i, 2) = rs收费细目!名称
                If Not IsNull(rs收费细目!产地) Then
                    .TextMatrix(i, 2) = .TextMatrix(i, 2) & "(" & rs收费细目!产地 & ")"
                End If
                If Not IsNull(rs收费细目!规格) Then
                    .TextMatrix(i, 2) = .TextMatrix(i, 2) & " " & rs收费细目!规格
                End If
                
                '计算单位:药嘱药品为门诊单位,非药嘱药品为售价单位
                .TextMatrix(i, 3) = Nvl(rsTmp!计算单位)
                '计价数量:药嘱药品为1,非药嘱药品为对应售价数
                .TextMatrix(i, 4) = FormatEx(rsTmp!数量, 5)
                
                '执行科室
                lng执行科室ID = Nvl(rsTmp!执行科室ID, 0)
                If rs收费细目!类别 = "4" And Nvl(rs收费细目!跟踪在用, 0) = 1 _
                    Or InStr(",5,6,7,", rs收费细目!类别) > 0 And InStr(",5,6,7,", rs诊疗项目!类别) = 0 Then
                    lng病人科室ID = UserInfo.部门ID
                    lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rs收费细目!类别, rs收费细目!ID, 4, lng病人科室ID, 0, 1, lng执行科室ID)
                End If
                
                '单价处理
                If InStr(",5,6,7,", rs收费细目!类别) > 0 Then
                    If Nvl(rs收费细目!是否变价, 0) = 1 Then
                        '求药品时价
                        If InStr(",5,6,7,", rs诊疗项目!类别) > 0 Then
                            '药嘱药品计算一个门诊包装的门诊时价
                            .TextMatrix(i, 5) = CalcDrugPrice(rs收费细目!ID, lng执行科室ID, Nvl(rsTmp!门诊包装, 1))
                            .TextMatrix(i, 5) = Format(Val(.TextMatrix(i, 5)) * Nvl(rsTmp!门诊包装, 0), "0.00000")
                        Else
                            '非药嘱药品计算相对售价数量的售价实价
                            .TextMatrix(i, 5) = Format(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, Nvl(rsTmp!数量, 0)), "0.00000")
                        End If
                    Else
                        '药嘱药品为门诊单价,非药药品为售价
                        .TextMatrix(i, 5) = Format(Nvl(rsTmp!单价), "0.00000")
                    End If
                ElseIf rs收费细目!类别 = "4" And Nvl(rs收费细目!跟踪在用, 0) = 1 And Nvl(rs收费细目!是否变价, 0) = 1 Then
                    '时价卫材的单价和药品一样计算
                    .TextMatrix(i, 5) = Format(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, Nvl(rsTmp!数量, 0)), "0.00000")
                Else
                    .TextMatrix(i, 5) = Format(Nvl(rsTmp!单价), "0.00000")
                End If

                '执行科室
                If lng执行科室ID <> 0 Then
                    .TextMatrix(i, 6) = Get部门名称(lng执行科室ID)
                End If
                
                '费用类型
                .TextMatrix(i, 7) = Nvl(rs收费细目!费用类型)
                
                '从属项目
                .TextMatrix(i, 8) = IIF(Nvl(rsTmp!从项, 0) = 0, "", "√")
                
                dblPrice = dblPrice + Format(Val(.TextMatrix(i, 4)) * Val(.TextMatrix(i, 5)), "0.00000")
                
                rsTmp.MoveNext
            Next
        End If
        
        '合计行
        If .Rows > 2 Then
            .Rows = .Rows + 1
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 3) = "合计"
            .Cell(flexcpAlignment, .Rows - 1, 0, .Rows - 1, 3) = 4
            .Cell(flexcpText, .Rows - 1, 4, .Rows - 1, 5) = Format(dblPrice, "0.00000")
            .Cell(flexcpAlignment, .Rows - 1, 4, .Rows - 1, 5) = 7
            .MergeCells = flexMergeFree
            .MergeRow(.Rows - 1) = True
        End If
        
        .Row = 1: .Col = 0
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    ShowPrice = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSendList(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱的发送记录
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strExe1 As String, strExe2 As String, strState As String
    Dim bln配方行 As Boolean, bln检验行 As Boolean
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .MergeCells = flexMergeNever
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSendList = True: Exit Function
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            bln配方行 = RowIs配方行(lngRow)
            bln检验行 = RowIs检验行(lngRow)
        End If
                
        strExe1 = "Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部份执行')"
        strExe2 = "Decode(Nvl(B.执行状态,0),0,'未执行',1,'执行完成',2,'拒绝执行',3,'正在执行')"
        strState = "Decode(A.记录性质,1,Decode(A.记录状态,0,'收费划价',1,'已收费',3,'已退费'),2,Decode(A.记录状态,0,'记帐划价',1,'已记帐',3,'已销帐'),'已计费')"
        
        '药嘱对应的药品计价按门诊包装显示,非药嘱对应的药品计价按零售单位显示
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            If Not RowIn一并给药(lngRow, lngBegin, lngEnd) Then lngBegin = lngRow
            '成药部份:填写了发送记录,但可能无对应费用(如自备药,但医嘱有规格)
            strSub = "Select A.*,B.门诊包装,B.门诊单位" & _
                " From 病人费用记录 A,药品规格 B" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL And A.收费类别 IN('5','6','7')" & _
                " And A.收费细目ID=B.药品ID And A.医嘱序号=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, "病人费用记录", "H病人费用记录")
            ElseIf MovedByDate(mvRegDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "病人费用记录", "H病人费用记录")
            End If
            
            strSQL = _
                " Select C.相关ID,C.标本部位,B.发送时间,B.NO,B.记录性质,A.收费细目ID," & _
                " Nvl(A.门诊单位,D.门诊单位) as 单位," & _
                " Nvl(A.数次/Nvl(A.门诊包装,1),B.发送数次/Nvl(D.剂量系数,1)/Nvl(D.门诊包装,1)) as 发送数次," & _
                " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID," & _
                " Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态,B.首次时间,B.末次时间," & _
                " Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费',1," & strState & ") as 计费状态," & _
                " B.发送人,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别" & _
                " From (" & strSub & ") A,病人医嘱发送 B,病人医嘱记录 C,药品规格 D" & _
                " Where B.医嘱ID=C.ID And C.收费细目ID=D.药品ID And C.ID=[1]" & _
                " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And A.医嘱序号(+)=B.医嘱ID"
            
            '在一并给药的首行才显示给药途径的发送
            If lngRow = lngBegin Then
                '给药途径部份:填写了发送记录(叮嘱无),但不一定有费用
                strSub = "Select A.*,B.门诊包装,B.门诊单位" & _
                    " From 病人费用记录 A,药品规格 B" & _
                    " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                    " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=[2]"
                If mblnMoved Then
                    strSub = Replace(strSub, "病人费用记录", "H病人费用记录")
                ElseIf MovedByDate(mvRegDate) Then
                    strSub = strSub & " Union ALL " & Replace(strSub, "病人费用记录", "H病人费用记录")
                End If
                    
                strSQL = strSQL & " Union ALL " & _
                    " Select C.相关ID,C.标本部位,B.发送时间,B.NO,B.记录性质,A.收费细目ID," & _
                    " Decode(Nvl(Instr('567',A.收费类别),0),0,D.计算单位,Nvl(A.门诊单位,E.门诊单位)) as 单位," & _
                    " Decode(Nvl(Instr('567',A.收费类别),0),0,B.发送数次," & _
                    "   Nvl(A.数次/Nvl(A.门诊包装,1),B.发送数次/Nvl(E.剂量系数,1)/Nvl(E.门诊包装,1))) as 发送数次," & _
                    " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID," & _
                    " Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态,B.首次时间," & _
                    " B.末次时间,Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费',1," & strState & ") as 计费状态," & _
                    " B.发送人,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别" & _
                    " From (" & strSub & ") A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 D,药品规格 E" & _
                    " Where B.医嘱ID=C.ID And C.诊疗项目ID=D.ID And C.收费细目ID=E.药品ID(+)" & _
                    " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And 0+A.医嘱序号(+)=B.医嘱ID And C.ID=[2]"
            End If
            
            If mblnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            End If
        Else
            '其它医嘱(包括配方及检查，手术一组医嘱):填写了发送记录(叮嘱无),但不一定有费用
            '中药自备药也是无对应费用(但医嘱有规格)
            strSub = _
                " Select A.*,B.门诊包装,B.门诊单位" & _
                " From 病人费用记录 A,药品规格 B" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=[1]"
            strSub = strSub & " Union ALL " & _
                " Select A.*,B.门诊包装,B.门诊单位" & _
                " From 病人费用记录 A,药品规格 B,病人医嘱记录 C" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=C.ID" & _
                " And C.相关ID=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, "病人费用记录", "H病人费用记录")
            ElseIf MovedByDate(mvRegDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "病人费用记录", "H病人费用记录")
            End If
            
            strSQL = _
                " Select * From 病人医嘱记录 Where ID=[1]" & _
                " Union ALL " & _
                " Select * From 病人医嘱记录 Where 相关ID=[1]"
            strSQL = _
                " Select C.相关ID,C.标本部位,B.发送时间,B.NO,B.记录性质,A.收费细目ID," & _
                " Decode(Nvl(Instr('567',A.收费类别),0),0,D.计算单位,Nvl(A.门诊单位,E.门诊单位)) as 单位," & _
                " Decode(Nvl(Instr('567',A.收费类别),0),0,B.发送数次," & _
                "   Nvl(Nvl(A.付数,1)*A.数次/Nvl(A.门诊包装,1),B.发送数次/Nvl(E.剂量系数,1)/Nvl(E.门诊包装,1))) as 发送数次," & _
                " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID," & _
                " Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态,B.首次时间,B.末次时间," & _
                " Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费',1," & strState & ") as 计费状态," & _
                " B.发送人,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别" & _
                " From (" & strSub & ") A,病人医嘱发送 B,(" & strSQL & ") C,诊疗项目目录 D,药品规格 E" & _
                " Where B.医嘱ID=C.ID And C.诊疗项目ID=D.ID And C.收费细目ID=E.药品ID(+)" & _
                " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And 0+A.医嘱序号(+)=B.医嘱ID"
            If mblnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            End If
        End If
        
        strSQL = "Select /*+ RULE */ A.发送序号,A.费用序号," & _
            " A.相关ID,A.诊疗类别,F.名称 as 类别名称,D.名称 as 诊疗项目,A.标本部位,A.发送时间,A.NO,A.记录性质," & _
            " Nvl(G.名称,B.名称)||Decode(B.产地,NULL,NULL,'('||B.产地||')')||Decode(B.规格,NULL,NULL,' '||B.规格) as 收费项目," & _
            " A.单位,A.发送数次 as 数量,C.名称 as 执行科室,A.执行状态,A.首次时间,A.末次时间,A.计费状态,A.发送人,A.发送号" & _
            " From (" & strSQL & ") A,收费项目目录 B,部门表 C,诊疗项目目录 D,诊疗项目类别 F,收费项目别名 G" & _
            " Where A.收费细目ID=B.ID(+) And A.执行部门ID=C.ID(+)" & _
            " And A.诊疗项目ID=D.ID And A.诊疗类别=F.编码" & _
            " And A.收费细目ID=G.收费细目ID(+) And G.码类(+)=1 And G.性质(+)=" & IIF(gbln商品名, 3, 1) & _
            " Order by A.发送号 Desc,A.诊疗类别,A.发送序号,A.费用序号"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, cs发送号) = Nvl(rsTmp!发送号, 0)
                .TextMatrix(i, cs发送时间) = Format(Nvl(rsTmp!发送时间), "MM-dd HH:mm")
                
                '发送医嘱
                If InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                    .TextMatrix(i, cs发送医嘱) = "药品医嘱-" & rsTmp!诊疗项目
                ElseIf rsTmp!诊疗类别 = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                    .TextMatrix(i, cs发送医嘱) = "给药途径-" & rsTmp!诊疗项目
                ElseIf rsTmp!诊疗类别 = "E" And (bln配方行 Or bln检验行) Then
                    If bln检验行 Then
                        .TextMatrix(i, cs发送医嘱) = "采集方法-" & rsTmp!诊疗项目
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        .TextMatrix(i, cs发送医嘱) = "中药煎法-" & rsTmp!诊疗项目
                    Else
                        .TextMatrix(i, cs发送医嘱) = "中药用法-" & rsTmp!诊疗项目
                    End If
                ElseIf Not IsNull(rsTmp!相关ID) Then
                    If rsTmp!诊疗类别 = "C" Then
                        .TextMatrix(i, cs发送医嘱) = "检验项目-" & rsTmp!诊疗项目
                    ElseIf rsTmp!诊疗类别 = "D" Then
                        .TextMatrix(i, cs发送医嘱) = "检查部位-" & Nvl(rsTmp!标本部位)
                    ElseIf rsTmp!诊疗类别 = "F" Then
                        .TextMatrix(i, cs发送医嘱) = "附加手术-" & rsTmp!诊疗项目
                    ElseIf rsTmp!诊疗类别 = "G" Then
                        .TextMatrix(i, cs发送医嘱) = "麻醉项目-" & rsTmp!诊疗项目
                    End If
                Else
                    .TextMatrix(i, cs发送医嘱) = rsTmp!类别名称 & "医嘱-" & rsTmp!诊疗项目
                End If
               
                .TextMatrix(i, cs单据号) = Nvl(rsTmp!NO)
                .TextMatrix(i, cs收费项目) = Nvl(rsTmp!收费项目)
                .TextMatrix(i, cs数次) = FormatEx(Nvl(rsTmp!数量), 5) & Nvl(rsTmp!单位)
                .TextMatrix(i, cs计费状态) = Nvl(rsTmp!计费状态)
                .TextMatrix(i, cs执行状态) = Nvl(rsTmp!执行状态)
                .TextMatrix(i, cs执行科室) = Nvl(rsTmp!执行科室)
                .TextMatrix(i, cs发送人) = Nvl(rsTmp!发送人)
                .TextMatrix(i, cs记录性质) = Nvl(rsTmp!记录性质)
                
                '已收费的划价单突出显示
                If .TextMatrix(i, cs计费状态) = "已缴费" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC00000 '深蓝
                ElseIf .TextMatrix(i, cs计费状态) = "已退费" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H808080 '灰色
                End If
                rsTmp.MoveNext
            Next
        End If
        
        .Row = 1: .Col = cs发送医嘱
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSendList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSignList(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱的签名记录
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .MergeCells = flexMergeNever
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSignList = True: Exit Function
        End If
        
        strSQL = "Select A.签名ID,A.操作类型,B.签名时间,B.签名人," & _
            " Decode(A.操作类型,1,'新开医嘱',4,'作废医嘱','其它操作') as 签名类型" & _
            " From 病人医嘱状态 A,医嘱签名记录 B Where A.医嘱ID=[1] And A.签名ID=B.ID Order by B.签名时间"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
            strSQL = Replace(strSQL, "医嘱签名记录", "H医嘱签名记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!签名ID)
                .TextMatrix(i, 0) = rsTmp!签名类型
                .Cell(flexcpData, i, 0) = Val(rsTmp!操作类型)
                .TextMatrix(i, 1) = Format(rsTmp!签名时间, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 2) = rsTmp!签名人
                Set .Cell(flexcpPicture, i, 0) = imgSign.ListImages(1).Picture
                rsTmp.MoveNext
            Next
        End If
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = 0
        .Row = 1
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSignList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowBillList() As Boolean
'功能：显示指定行的医嘱发送可以打印的诊疗单据在菜单上
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objMenu As Menu, lng医嘱ID As Long
    
    For i = mfrmParent.mnuReportClinic.UBound To 0 Step -1
        mfrmParent.mnuReportClinic(i).Tag = ""
        If i = 0 Then
            mfrmParent.mnuReportClinic(i).Caption = "<无可用单据>"
        Else
            Unload mfrmParent.mnuReportClinic(i)
        End If
    Next
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Then
        ShowBillList = True: Exit Function
    ElseIf tabAppend.SelectedItem.Index <> 2 Then
        ShowBillList = True: Exit Function
    ElseIf vsAppend.TextMatrix(vsAppend.Row, cs发送号) = "" Then
        ShowBillList = True: Exit Function
    End If
        
    On Error GoTo errH
        
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID)) <> 0 Then
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID))
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    
    With vsAppend
        strSQL = "Select Distinct D.编号,D.名称,D.说明" & _
            " From 病人医嘱发送 A,病人医嘱记录 B,诊疗单据应用 C,病历文件目录 D" & _
            " Where A.发送号=[1] And A.NO=[2]" & _
            " And A.医嘱ID=B.ID And B.诊疗项目ID=C.诊疗项目ID" & _
            " And C.应用场合=1 And C.病历文件ID=D.ID And D.种类=5" & _
            " And D.前提 IN(1,3) And D.书写 IN(1,2)" & _
            " And (B.ID=[3] Or B.相关ID=[3])" & _
            " Order by D.编号"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(.Row, cs发送号)), .TextMatrix(.Row, cs单据号), lng医嘱ID)
    End With
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If i > 1 Then Load mfrmParent.mnuReportClinic(mfrmParent.mnuReportClinic.UBound + 1)
            Set objMenu = mfrmParent.mnuReportClinic(mfrmParent.mnuReportClinic.UBound)
            objMenu.Caption = rsTmp!名称
            If i <= 10 Then
                objMenu.Caption = objMenu.Caption & "(&" & i - 1 & ")"
            ElseIf i <= 36 Then
                objMenu.Caption = objMenu.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
            End If
            objMenu.Tag = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" '对应的自定义报表编号
            'If i > 1 Then objMenu.Enabled = False '一个项目只能设置一个诊疗单据
            rsTmp.MoveNext
        Next
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAppend_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    If NewCol >= vsAppend.FixedCols And NewRow >= vsAppend.FixedRows Then
        If vsAppend.Redraw <> flexRDNone Then
            '为了外部系统调用增加，By：赵彤宇
            On Error Resume Next
            If mfrmParent.mnuReportClinic.UBound < 0 Then Exit Sub
            On Error GoTo 0
        
            Call ShowBillList '显示可打印的诊疗单据
        End If
    End If
End Sub

Private Sub vsAppend_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    
    With vsAppend
        If Button = 2 And tabAppend.SelectedItem.Index = 2 Then
            If Between(.MouseRow, .FixedRows, .Rows - 1) Then
                If Between(.MouseCol, .FixedCols, .Cols - 1) Then
                    If mfrmParent.mnuReportItem(mnu打印诊疗单据).Enabled Then
                        PopupMenu mfrmParent.mnuReportItem(mnu打印诊疗单据), 2
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAppend_GotFocus()
    vsAppend.BackColorSel = &HFFCC99
End Sub

Private Sub vsAppend_LostFocus()
    vsAppend.BackColorSel = &HFFEBD7
End Sub

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub ShowTotalMoney()
'功能：医嘱总金额的提示
'说明：由于药品时价，和给药途径，中药煎法用法等，新开医嘱不一定准确
    Dim rsMoney As New ADODB.Recordset, strSQL As String
    Dim cur应收 As Currency, cur实收 As Currency
    Dim cur药品应收 As Currency, cur药品实收 As Currency
    Dim cur新开 As Currency, cur药品新开 As Currency
    Dim cur预交 As Currency
    
    On Error GoTo errH
    
    strSQL = _
        " Select /*+ RULE */ Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
        " Sum(Decode(Instr('567',A.收费类别),0,0,A.应收金额)) as 药品应收," & _
        " Sum(Decode(Instr('567',A.收费类别),0,0,A.实收金额)) as 药品实收" & _
        " From 病人费用记录 A,病人医嘱发送 B,病人医嘱记录 C" & _
        " Where A.医嘱序号=B.医嘱ID And B.医嘱ID=C.ID" & _
        " And C.病人ID+0=[1] And C.挂号单=[2]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
    ElseIf MovedByDate(mvRegDate) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, "病人费用记录", "H病人费用记录")
        strSQL = "Select Sum(应收金额) as 应收金额,Sum(实收金额) as 实收金额," & _
            " Sum(药品应收) as 药品应收,Sum(药品实收) as 药品实收 From (" & strSQL & ")"
    End If
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mstr挂号单)
    If Not rsMoney.EOF Then
        cur应收 = Nvl(rsMoney!应收金额, 0)
        cur实收 = Nvl(rsMoney!实收金额, 0)
        cur药品应收 = Nvl(rsMoney!药品应收, 0)
        cur药品实收 = Nvl(rsMoney!药品实收, 0)
    End If
        
    '时价药品取"指导零售价"
    strSQL = _
        "Select Sum(Round(金额," & gbytDec & ")) As 金额,Sum(Round(药品金额," & gbytDec & ")) As 药品金额" & _
        " From (Select A.总给予量*Decode(I.是否变价,1,S.指导零售价,P.现价) As 金额," & _
        "              A.总给予量*Decode(I.是否变价,1,S.指导零售价,P.现价) As 药品金额" & _
        "       From 病人医嘱记录 A,收费项目目录 I,收费价目 P,药品规格 S" & _
        "       Where A.收费细目ID=I.ID And I.ID=P.收费细目ID And I.ID=S.药品ID" & _
        "             And (Sysdate Between P.执行日期 And P.终止日期 Or Sysdate>=P.执行日期 And P.终止日期 is Null)" & _
        "             And A.医嘱状态=1 And A.诊疗类别 In ('5','6')" & _
        "             And A.病人ID+0=[1] And A.挂号单=[2]" & _
        "       Union All" & _
        "       Select A.总给予量*A.单次用量/S.剂量系数*Decode(I.是否变价,1,S.指导零售价,P.现价) As 金额," & _
        "              A.总给予量*A.单次用量/S.剂量系数*Decode(I.是否变价,1,S.指导零售价,P.现价) As 药品金额" & _
        "       From 病人医嘱记录 A,收费项目目录 I,收费价目 P,药品规格 S" & _
        "       Where A.收费细目ID=I.ID And I.ID=P.收费细目ID And I.ID=S.药品ID" & _
        "             And (Sysdate Between P.执行日期 And P.终止日期 Or Sysdate>=P.执行日期 And P.终止日期 is Null)" & _
        "             And A.医嘱状态=1 And A.诊疗类别='7'" & _
        "             And A.病人ID+0=[1] And A.挂号单=[2]" & _
        "       Union All" & _
        "       Select Nvl(A.总给予量,A.频率次数)*R.收费数量*P.现价 As 金额,0 as 药品金额" & _
        "       From 病人医嘱记录 A,诊疗收费关系 R,收费项目目录 I,收费价目 P" & _
        "       Where A.诊疗项目ID=R.诊疗项目ID And I.ID=R.收费项目ID And I.ID=P.收费细目ID" & _
        "             And (Sysdate Between P.执行日期 And P.终止日期 Or Sysdate>=P.执行日期 And P.终止日期 is Null)" & _
        "             And Nvl(A.计价特性,0)=0 And A.医嘱状态=1 And A.诊疗类别 Not In ('5','6','7')" & _
        "             And A.病人ID+0=[1] And A.挂号单=[2]) A"
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mstr挂号单)
    If Not rsMoney.EOF Then
        cur新开 = Nvl(rsMoney!金额, 0)
        cur药品新开 = Nvl(rsMoney!药品金额, 0)
    End If
    
    strSQL = "Select Nvl(预交余额,0)-Nvl(费用余额,0) as 金额 From 病人余额 Where 性质=1 And 病人ID=[1]"
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID)
    If Not rsMoney.EOF Then cur预交 = Nvl(rsMoney!金额, 0)
    
    mfrmParent.stbThis.Panels(2).Text = _
        "医嘱已发送应收:" & FormatEx(cur应收, gbytDec) & "(药" & FormatEx(cur药品应收, gbytDec) & ")," & _
        "实收:" & FormatEx(cur实收, gbytDec) & "(药" & FormatEx(cur药品实收, gbytDec) & ")" & _
        "  新开约:" & FormatEx(cur新开, gbytDec) & "(药" & FormatEx(cur药品新开, gbytDec) & ")" & _
        IIF(cur预交 = 0, "", "  预交:" & FormatEx(cur预交, "0.00"))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
