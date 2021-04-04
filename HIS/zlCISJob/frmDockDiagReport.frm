VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockDiagReport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Height          =   4380
      Left            =   75
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
      FormatString    =   $"frmDockDiagReport.frx":0000
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
               Picture         =   "frmDockDiagReport.frx":009B
               Key             =   "未填"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockDiagReport.frx":05B5
               Key             =   "已填"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDockDiagReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Activate() '自已激活时
Public Event RequestRefresh(ByVal RefreshNotify As Boolean) '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字

Private mMainPrivs As String '主模块权限
Private mfrmParent As Object
Private mcbsMain As CommandBars
Private mint场合 As Integer
Private mbln护士站 As Boolean
Private mlng病人ID As Long
Private mvar就诊ID As Variant
Private mblnMoved As Boolean '病人住院数据是否已转出

Private Enum PATI_TYPE
    '住院
    pt在院 = 0
    pt预出 = 1
    pt出院 = 2
    pt待诊 = 3 '医生站:待会诊病人(在院)
    pt已诊 = 4 '医生站:已会诊病人
    '门诊
    pt禁止 = 0
    pt允许 = 1
End Enum
Private mint类型 As PATI_TYPE

'存放当前可用单据列表
Private Type TYPE_Bill
    ID As Long
    名称 As String
End Type
Private marrBill() As TYPE_Bill

'列常量
Private Enum BILL_COL
    COL_F申请 = 0 '标志列
    COL_F报告 = 1
    COL_NO = 2 '可见列
    COL_医嘱内容 = 3
    COL_单据 = 4
    COL_申请人 = 5
    COL_申请时间 = 6
    COL_发送时间 = 7
    COL_报告人 = 8
    COL_报告时间 = 9
    COL_医嘱ID = 10 '隐藏列
    COL_诊疗项目ID = 11
    COL_单据ID = 12
    COL_编号 = 13
    COL_申请项 = 14
    COL_申请ID = 15
    COL_报告项 = 16
    COL_报告ID = 17
    COL_记录性质 = 18
End Enum

Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal var就诊ID As Variant, ByVal int类型 As Integer, Optional ByVal blnMoved As Boolean)
'功能：刷新或清除单据清单
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    mlng病人ID = lng病人ID: mvar就诊ID = var就诊ID
    mint类型 = int类型: mblnMoved = blnMoved
        
    vsBill.Rows = vsBill.FixedRows
    vsBill.Rows = vsBill.FixedRows + 1
    
    Call LoadBillList
    If mlng病人ID <> 0 Then
        Call LoadReport
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As CommandBars, ByVal int场合 As Integer, Optional ByVal bln护士站 As Boolean)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    mint场合 = int场合: mbln护士站 = bln护士站
    Set mfrmParent = frmParent
    Set mcbsMain = cbsMain
    cbsMain.Icons = frmPubIcons.imgPublic.Icons
    
    '辅诊菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "辅诊(&E)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "新增申请单(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改申请单(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除申请单(&D)")
                
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "填写报告单(&W)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "查阅申请单(&T)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "查阅报告单(&R)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "打印通知单(&I)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend, "打印报告单(&P)")
        
        If Not mbln护士站 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片处理(&V)"): objControl.BeginGroup = True
        End If
    End With

    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    Set objBar = cbsMain(2)
    For Each objControl In objBar.Controls '先求出前面的最后一个Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    With objBar.Controls
        'Set objControl = .Find(, conMenu_File_Preview) '从预览按钮之后开始加入
        Set objPopup = .Add(xtpControlPopup, conMenu_Edit_NewItem, "申请", objControl.Index + 1): objPopup.BeginGroup = True
        objPopup.ID = conMenu_Edit_NewItem
        objPopup.IconId = conMenu_Edit_NewItem
        objPopup.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改", objPopup.Index + 1): objControl.ToolTipText = "修改申请单"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除", objControl.Index + 1): objControl.ToolTipText = "删除申请单"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "报告", objControl.Index + 1): objControl.BeginGroup = True
        If Not mbln护士站 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片", objControl.Index + 1): objControl.BeginGroup = True
        End If
    End With
    
    '命令的快键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新增申请单
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改申请单
        .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除申请单
        .Add FCONTROL, vbKeyR, conMenu_Edit_Append '填写报告单
        .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '观片处理
    End With

    '设置不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
    End With
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
'功能：根据权限、当前病人或数据情况，设置功能或可见和可用性
'  1.无病人的情况
'  2.病人已出院(就诊)的情况
'  3.无数据的情况
    Dim blnBill As Boolean, blnEnabled As Boolean
            
    If vsBill.Redraw = flexRDNone Then Exit Sub
        
    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    '辅诊操作部份
    '------------------------------------------------------------------------------
    '总的判断:无病人不允许任何操作
    If Between(Control.ID, conMenu_Edit_NewItem, conMenu_Edit_NewItem + 999) Then
        Control.Enabled = mlng病人ID <> 0
        If Not Control.Enabled Then Exit Sub
    End If
    
    '辅诊部份
    '------------------------------------------------------------------------------
    blnBill = Val(vsBill.TextMatrix(vsBill.Row, COL_医嘱ID)) <> 0
    blnEnabled = mint场合 = 1 And mint类型 = pt允许 Or mint场合 = 2 And (mint类型 = pt在院 Or mint类型 = pt待诊)
    Select Case Control.ID
        Case conMenu_Edit_NewItem
            Control.Enabled = blnEnabled And UBound(marrBill) >= 1
        Case conMenu_Edit_NewItem * 100# + 1 To conMenu_Edit_NewItem * 100# + 200 '新增申请单
            Control.Enabled = blnEnabled
        Case conMenu_Edit_Audit '查阅申请单
            Control.Enabled = blnBill
        Case conMenu_Edit_Modify '修改申请单
            Control.Enabled = blnBill And blnEnabled
        Case conMenu_Edit_Delete '删除申请单
            Control.Enabled = blnBill And blnEnabled
        Case conMenu_Edit_Adjust '打印通知单
            Control.Enabled = blnBill And blnEnabled
        Case conMenu_Edit_SendBack '查阅报告单
            Control.Enabled = blnBill
        Case conMenu_Edit_Append '填写报告单
            Control.Enabled = blnBill And blnEnabled
        Case conMenu_Edit_Compend '打印报告单
            Control.Enabled = blnBill And blnEnabled
        Case conMenu_Edit_MarkMap '观片处理
            Control.Enabled = blnBill And blnEnabled
            If Control.Enabled Then
                Control.Enabled = (vsBill.Cell(flexcpData, vsBill.Row, COL_诊疗项目ID) = "D" And vsBill.Cell(flexcpData, vsBill.Row, COL_报告ID) = 1)
            End If
    End Select
    
    '其它部份
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = blnBill
    End Select
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置菜单和工具栏的可见状态
    Dim blnVisible As Boolean, strItem As String

    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Category = "已判断" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_NewItem '新增申请单
            If InStr(GetInsidePrivs(p辅诊记录管理), ";申请填写;") = 0 Then blnVisible = False
        Case conMenu_Edit_Audit '查阅申请单
            If InStr(GetInsidePrivs(p辅诊记录管理), ";申请填写;") = 0 Then blnVisible = False
        Case conMenu_Edit_Modify '修改申请单
            If InStr(GetInsidePrivs(p辅诊记录管理), ";申请填写;") = 0 Then blnVisible = False
        Case conMenu_Edit_Delete '删除申请单
            If InStr(GetInsidePrivs(p辅诊记录管理), ";申请填写;") = 0 Then blnVisible = False
        Case conMenu_Edit_Adjust '打印通知单
            If InStr(GetInsidePrivs(p辅诊记录管理), ";申请填写;") = 0 Then blnVisible = False
        Case conMenu_Edit_SendBack '查阅报告单
            If InStr(GetInsidePrivs(p辅诊记录管理), ";报告查阅;") = 0 Then blnVisible = False
        Case conMenu_Edit_Append '填写报告单
            If InStr(GetInsidePrivs(p辅诊记录管理), ";报告编辑;") = 0 Then blnVisible = False
        Case conMenu_Edit_Compend '打印报告单
            If InStr(GetInsidePrivs(p辅诊记录管理), ";报告打印;") = 0 Then blnVisible = False
        Case conMenu_Edit_MarkMap '观片处理
            If InStr(GetInsidePrivs(p辅诊记录管理), ";观片处理;") = 0 Then blnVisible = False
    End Select
    
    Control.Visible = blnVisible
    Control.Category = "已判断"
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    Dim objControl As CommandBarControl
    Dim vBill As TYPE_Bill, i As Long
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_NewItem '新增申请单
        With CommandBar.Controls
            .DeleteAll
            For i = 1 To UBound(marrBill)
                vBill = marrBill(i)
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 100# + i, vBill.名称)
                If i <= 10 Then
                    objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                ElseIf i <= 36 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                End If
                objControl.Parameter = vBill.ID
            Next
        End With
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet '打印设置
            Call zlPrintSet
        Case conMenu_File_Preview '预览辅诊清单
            Call OutputList(2)
        Case conMenu_File_Print '打印辅诊清单
            Call OutputList(1)
        Case conMenu_File_Excel '输出辅诊清单
            Call OutputList(3)
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Tool_Reference_2 '诊疗措拖参考
            Call zlItemRef
        '------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem * 100# + 1 To conMenu_Edit_NewItem * 100# + 200 '新增申请单
            Call FuncAddRequest(Val(Control.Parameter))
        Case conMenu_Edit_Audit '查阅申请单
            Call FuncWriteRequest(True)
        Case conMenu_Edit_Modify '修改申请单
            Call FuncWriteRequest(False)
        Case conMenu_Edit_Delete '删除申请单
            Call FuncDeleteRequest
        Case conMenu_Edit_Adjust '打印通知单
            Call FuncPrintRequest
        Case conMenu_Edit_SendBack '查阅报告单
            Call FuncWriteReport(True)
        Case conMenu_Edit_Append '填写报告单
            Call FuncWriteReport(False)
        Case conMenu_Edit_Compend '打印报告单
            Call FuncPrintReport
        Case conMenu_Edit_MarkMap '观片处理
            Call ViewImage(Val(vsBill.TextMatrix(vsBill.Row, COL_医嘱ID)), mfrmParent, mblnMoved)
    End Select
End Sub

Private Sub InitBillTable()
'功能：初始化单据清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "单据号,810,1;医嘱内容,3000,1;单据,1800,1;申请人,850,1;" & _
        "申请时间,1080,1;发送时间,1080,1;报告人,850,1;报告时间,1080,1;" & _
        "医嘱ID;诊疗项目ID;单据ID;编号;申请项;申请ID;报告项;报告ID;记录性质"
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

Private Sub Form_Load()
    Call InitBillTable
    Call RestoreWinState(Me, App.ProductName)
    mMainPrivs = gstrPrivs
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsBill.Left = 0
    vsBill.Top = 0
    vsBill.Width = Me.ScaleWidth
    vsBill.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub FuncPrintRequest()
'功能：打印通知单
    Dim strBill As String
    
    If mlng病人ID = 0 Then Exit Sub
    
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
        '如果未发送则不允许
        If Len(.TextMatrix(.Row, COL_NO)) = 0 Then
            MsgBox "该申请还未发送，不能打印通知单。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strBill = "ZLCISBILL" & Format(.TextMatrix(.Row, COL_编号), "00000") & "-1"
        If ReportPrintSet(gcnOracle, glngSys, strBill, mfrmParent) Then
            Call ReportOpen(gcnOracle, glngSys, strBill, mfrmParent, "NO=" & .TextMatrix(.Row, COL_NO), "性质=" & Val(.TextMatrix(.Row, COL_记录性质)), 2)
        End If
    End With
End Sub

Private Sub FuncPrintReport(Optional ByVal PrtMode As Integer = 2)
'功能：打印报告单
    Dim strBill As String
    
    If mlng病人ID = 0 Then Exit Sub
    
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
        
'        strBill = "ZLCISBILL" & Format(.TextMatrix(.Row, COL_编号), "00000") & "-2"
'        If ReportPrintSet(gcnOracle, glngSys, strBill, mfrmParent) Then
'            Call ReportOpen(gcnOracle, glngSys, strBill, mfrmParent, "NO=" & .TextMatrix(.Row, COL_NO), "性质=" & Val(.TextMatrix(.Row, COL_记录性质)), 2)
'        End If
        Call PrintDiagRpt_New(Val(.TextMatrix(.Row, COL_报告ID)), mfrmParent, PrtMode, , mblnMoved)
    End With
End Sub

Private Sub FuncAddRequest(ByVal lng单据ID As Long)
'功能：新增申请单
    If mlng病人ID = 0 Then Exit Sub
    If lng单据ID = 0 Then Exit Sub
    
    If mblnMoved Then
        MsgBox "病人的本次就诊数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '调用接口
    Call AddRequest(mfrmParent, mlng病人ID, mvar就诊ID, lng单据ID, mbln护士站)
    If True Then
        Call LoadReport
    End If
End Sub

Private Sub FuncWriteRequest(ByVal blnReadOnly As Boolean)
'功能：填写辅诊申请
    If mlng病人ID = 0 Then Exit Sub
    With vsBill
        If Val(.TextMatrix(.Row, COL_医嘱ID)) = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, COL_申请项)) = 0 Then
            MsgBox "该单据不需要填写申请。", vbInformation, gstrSysName
            Exit Sub
        End If
        If Not blnReadOnly Then
            If .TextMatrix(.Row, COL_NO) <> "" Then
                MsgBox "该医嘱已经发送，不能再填写申请。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If mblnMoved Then
                MsgBox "病人的本次就诊数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                    "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '填写申请:医嘱ID,单据ID,申请ID,医嘱内容
        Call EditRequest(Me, Val(.TextMatrix(.Row, COL_医嘱ID)), Val(.TextMatrix(.Row, COL_单据ID)), _
            Val(.TextMatrix(.Row, COL_申请ID)), .TextMatrix(.Row, COL_医嘱内容), blnReadOnly, , , mblnMoved)
        If True Then
            Call LoadReport
        End If
    End With
End Sub

Private Sub FuncWriteReport(ByVal blnReadOnly As Boolean)
'功能：填写辅诊报告
    If mlng病人ID = 0 Then Exit Sub
    
    With vsBill
        If Val(.TextMatrix(.Row, COL_医嘱ID)) = 0 Then Exit Sub
        
        If Val(.TextMatrix(.Row, COL_报告项)) = 0 Then
            MsgBox "该单据不需要填写报告。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .TextMatrix(.Row, COL_NO) = "" Then
            MsgBox "该医嘱尚未发送，请先发送医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
'        If .RowData(.Row) > 0 And .RowData(.Row) < 6 Then
        
        If blnReadOnly Then
            If .TextMatrix(.Row, COL_报告ID) = 0 Then
                MsgBox "该医嘱尚未报告，不能查阅。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        If Not blnReadOnly Then
            If .Cell(flexcpData, .Row, COL_报告ID) <> 1 Then
                MsgBox "该医嘱的报告尚未审核，不能编辑。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If mblnMoved Then
                MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                    "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '填写报告:NO,记录性质,单据ID,报告ID,医嘱内容
        Call EditReport(Me, .TextMatrix(.Row, COL_NO), Val(.TextMatrix(.Row, COL_记录性质)), _
            Val(.TextMatrix(.Row, COL_单据ID)), Val(.TextMatrix(.Row, COL_报告ID)), _
            .TextMatrix(.Row, COL_医嘱内容), blnReadOnly Or .Cell(flexcpData, .Row, COL_报告ID) = 1, lng医嘱ID:=.TextMatrix(.Row, COL_医嘱ID))
        If True Then
            Call LoadReport
        End If
    End With
End Sub

Private Sub FuncDeleteRequest()
'功能：删除当前申请单
    Dim strSQL As String, lngRow As Long
        
    If mlng病人ID = 0 Then Exit Sub
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
        
        If mblnMoved Then
            MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("确实要删除申请单[" & .TextMatrix(.Row, COL_单据) & "]吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '已校对对的在过程中检查；对检验组合,注意这里的医嘱ID正好是采集方法的ID
        strSQL = "zl_病人医嘱记录_Delete(" & Val(.TextMatrix(.Row, COL_医嘱ID)) & ",1)"
    End With
    
    '删除申请单
    On Error GoTo errH
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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

Public Sub zlItemRef()
'功能：调用诊疗参考
    Dim lng诊疗项目ID As Long
    
    lng诊疗项目ID = Val(vsBill.TextMatrix(vsBill.Row, COL_诊疗项目ID))
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
    objOut.Title.Text = "病人单据清单"
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
    Set objOut.Body = vsBill
    
    '输出
    vsBill.Redraw = False
    lngRow = vsBill.Row: lngCol = vsBill.Col
    
    strWidth = ""
    For i = 0 To vsBill.FixedCols - 1
        strWidth = strWidth & "," & vsBill.ColWidth(i)
        vsBill.ColWidth(i) = 0
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsBill.FixedCols - 1
        vsBill.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    
    vsBill.Row = lngRow: vsBill.Col = lngCol
    vsBill.Redraw = True
End Sub

Private Function GetPatiInfo() As String
'功能：读取病人信息串(用于打印)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If mint场合 = 1 Then
        '执行部门(号别科室)即病人科室
        strSQL = "Select B.姓名,B.性别,B.年龄,B.门诊号," & _
            " B.险类,B.就诊诊室,C.名称 as 执行部门,A.执行部门ID,A.登记时间" & _
            " From 病人挂号记录 A,病人信息 B,部门表 C" & _
            " Where A.NO=[2] And A.病人ID+0=[1]" & _
            " And A.病人ID=B.病人ID And A.执行部门ID=C.ID"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mvar就诊ID)
        
        GetPatiInfo = _
            "姓名：" & rsTmp!姓名 & " 性别：" & Nvl(rsTmp!性别) & _
            " 年龄：" & Nvl(rsTmp!年龄) & " 门诊号：" & Nvl(rsTmp!门诊号) & _
            " 挂号：" & Format(rsTmp!登记时间, "MM-dd HH:mm") & _
            " 科室：" & rsTmp!执行部门 & " 诊室：" & Nvl(rsTmp!就诊诊室)
    ElseIf mint场合 = 2 Then
        strSQL = "Select B.姓名,B.性别,B.年龄,B.住院号," & _
            " B.险类,C.名称 as 科室,A.入院日期,A.出院日期" & _
            " From 病案主页 A,病人信息 B,部门表 C" & _
            " Where A.主页ID=[2] And A.病人ID=[1]" & _
            " And A.病人ID=B.病人ID And A.出院科室ID=C.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mvar就诊ID)
        
        GetPatiInfo = _
            "姓名：" & rsTmp!姓名 & " 性别：" & Nvl(rsTmp!性别) & _
            " 年龄：" & Nvl(rsTmp!年龄) & " 住院号：" & Nvl(rsTmp!住院号) & _
            " 科室：" & rsTmp!科室 & " 入院：" & Format(rsTmp!入院日期, "MM-dd HH:mm") & _
            IIf(Not IsNull(rsTmp!出院日期), " 出院：" & Format(rsTmp!出院日期, "MM-dd HH:mm"), "")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = COL_医嘱内容 Then
        vsBill.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsBill.TextMatrix(vsBill.FixedRows - 1, Col) & "A")
        If vsBill.ColWidth(Col) < lngW Then
            vsBill.ColWidth(Col) = lngW
        ElseIf vsBill.ColWidth(Col) > vsBill.Width * 0.5 Then
            vsBill.ColWidth(Col) = vsBill.Width * 0.5
        End If
    End If
End Sub

Private Sub vsBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsBill.FixedCols - 1 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsBill_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim vRect As RECT
    With vsBill
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
        End If
        Done = True
    End With
End Sub

Private Sub vsBill_GotFocus()
    vsBill.BackColorSel = &HFFCC99
End Sub

Private Sub vsBill_LostFocus()
    vsBill.BackColorSel = &HFFEBD7
End Sub

Private Function LoadReport() As Boolean
'功能：根据当前病人医嘱读取可以填写的申请单或报告单
    Dim rsBill As New ADODB.Recordset
    Dim strSQL As String, strBill As String
    Dim strKey As String, lngPreRow As Long
    Dim lngRow As Long, blnRemove As Boolean, i As Long
    
    If mlng病人ID = 0 Then Exit Function
    
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
        " Where A.诊疗项目ID=B.诊疗项目ID And B.应用场合=[3]" & _
        " And B.病历文件ID=C.病历文件ID And B.病历文件ID=D.ID" & _
        IIf(mint场合 = 1, " And A.病人ID+0=[1] And A.挂号单=[2]", " And A.病人ID=[1] And A.主页ID=[2]") & _
        " And (A.诊疗类别 Not IN('5','6','7') And A.相关ID is NULL" & _
        "   Or A.诊疗类别='C' And A.相关ID is Not NULL)" & _
        " Group by A.ID,B.病历文件ID,D.编号,D.名称,D.说明"
    
    '与药品相关的医嘱以及采集方法
    strSQL = "Select Distinct 相关ID From 病人医嘱记录" & _
        " Where 病人ID=[1] And " & IIf(mint场合 = 1, "挂号单", "主页ID") & "=[2]" & _
        " And (诊疗类别 IN('5','6','7') Or 诊疗类别='C' And 相关ID is Not NULL)"
        
    '医嘱对应的单据清单(包括待安排医嘱,不含不发送的叮嘱),至少包含一种单据附项
    '未发送的医嘱显示一条,已发送的一次发送显示一条(包括只有申请项的)
    strSQL = _
        " Select A.ID,A.相关ID,A.诊疗类别,A.诊疗项目ID,A.医嘱内容,A.标本部位," & _
        " B.发送时间,B.NO,B.记录性质,A.申请ID,B.报告ID,C.编号,C.名称,C.单据ID,C.申请项,C.报告项," & _
        " X.书写人 as 申请人,X.书写日期 as 申请时间," & _
        " Y.书写人 as 报告人,Y.书写日期 as 报告时间,Nvl(B.执行过程,0) As 执行过程,Nvl(B.执行状态,0) As 执行状态" & _
        " From 病人医嘱记录 A,病人医嘱发送 B,(" & strBill & ") C,病人病历记录 X,病人病历记录 Y" & _
        " Where " & IIf(mint场合 = 1, " A.病人ID+0=[1] And A.挂号单=[2]", " A.病人ID=[1] And A.主页ID=[2]") & _
        " And (A.诊疗类别 Not IN('5','6','7') And A.相关ID is NULL" & _
        "   Or A.诊疗类别='C' And A.相关ID is Not NULL)" & _
        " And A.ID Not IN(" & strSQL & ") And A.医嘱状态<>4 And Nvl(A.执行性质,0)<>0" & _
        " And A.ID=B.医嘱ID(+) And A.ID=C.医嘱ID And (C.申请项=1 Or C.报告项=1)" & _
        " And A.申请ID=X.ID(+) And B.报告ID=Y.ID(+)" & _
        " Order by Nvl(B.发送时间,A.开嘱时间) Desc,A.序号"
            
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
    End If
            
    '医嘱内容,NO,单据,申请人,申请时间,发送时间,报告人,报告时间
    '医嘱ID;诊疗项目ID;单据ID;编号;申请项;申请ID;报告项;报告ID;记录性质
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mvar就诊ID, mint场合)
    With vsBill
        .Rows = .FixedRows + rsBill.RecordCount
        lngRow = .FixedRows
        For i = 1 To rsBill.RecordCount
            .TextMatrix(lngRow, COL_医嘱内容) = rsBill!医嘱内容
                .Cell(flexcpData, lngRow, COL_医嘱内容) = Nvl(rsBill!标本部位) '检验标本
            .TextMatrix(lngRow, COL_NO) = Nvl(rsBill!NO)
            .TextMatrix(lngRow, COL_单据) = rsBill!名称
            .TextMatrix(lngRow, COL_申请人) = Nvl(rsBill!申请人)
            .TextMatrix(lngRow, COL_申请时间) = Format(Nvl(rsBill!申请时间), "MM-dd HH:mm")
            .TextMatrix(lngRow, COL_发送时间) = Format(Nvl(rsBill!发送时间), "MM-dd HH:mm")
            .TextMatrix(lngRow, COL_报告人) = Nvl(rsBill!报告人)
            .TextMatrix(lngRow, COL_报告时间) = Format(Nvl(rsBill!报告时间), "MM-dd HH:mm")
            .TextMatrix(lngRow, COL_医嘱ID) = rsBill!ID '隐藏列
            .TextMatrix(lngRow, COL_诊疗项目ID) = rsBill!诊疗项目ID
                .Cell(flexcpData, lngRow, COL_诊疗项目ID) = Nvl(rsBill!诊疗类别)
            .TextMatrix(lngRow, COL_单据ID) = rsBill!单据ID
            .TextMatrix(lngRow, COL_编号) = rsBill!编号
            .TextMatrix(lngRow, COL_申请项) = Nvl(rsBill!申请项, 0)
            .TextMatrix(lngRow, COL_申请ID) = Nvl(rsBill!申请ID, 0)
            .TextMatrix(lngRow, COL_报告项) = Nvl(rsBill!报告项, 0)
            .TextMatrix(lngRow, COL_报告ID) = Nvl(rsBill!报告ID, 0)
                .Cell(flexcpData, lngRow, COL_报告ID) = Nvl(rsBill!执行状态, 0)
            .TextMatrix(lngRow, COL_记录性质) = Nvl(rsBill!记录性质)
            .RowData(lngRow) = Nvl(rsBill!执行过程, 0)
            '已审核的报告，字体加粗
            .Cell(flexcpFontBold, lngRow, 1, lngRow, .Cols - 1) = Not (.Cell(flexcpData, lngRow, COL_报告ID) <> 1 Or _
                .TextMatrix(lngRow, COL_报告ID) = 0)
            
            '申请与报告的标识
            If rsBill!申请项 = 1 Then
                If Not IsNull(rsBill!申请ID) Then
                    Set .Cell(flexcpPicture, lngRow, COL_F申请) = imgFlag.ListImages("已填").Picture
                Else
                    Set .Cell(flexcpPicture, lngRow, COL_F申请) = imgFlag.ListImages("未填").Picture
                End If
            End If
            If rsBill!报告项 = 1 Then
                If Not IsNull(rsBill!报告ID) Then
                    Set .Cell(flexcpPicture, lngRow, COL_F报告) = imgFlag.ListImages("已填").Picture
                Else
                    Set .Cell(flexcpPicture, lngRow, COL_F报告) = imgFlag.ListImages("未填").Picture
                End If
            End If
            
            '删除其它一并采集的检验项目行
            blnRemove = False
            If rsBill!诊疗类别 = "C" And Not IsNull(rsBill!相关ID) Then
                .TextMatrix(lngRow, COL_医嘱ID) = rsBill!相关ID '一并采集的记录为相关ID
                If Val(.TextMatrix(lngRow - 1, COL_医嘱ID)) = rsBill!相关ID Then
                    '组合医嘱内容
                    .TextMatrix(lngRow - 1, COL_医嘱内容) = Replace(.TextMatrix(lngRow - 1, COL_医嘱内容), "(" & .Cell(flexcpData, lngRow - 1, COL_医嘱内容) & ")", "")
                    .TextMatrix(lngRow - 1, COL_医嘱内容) = .TextMatrix(lngRow - 1, COL_医嘱内容) & "," & .TextMatrix(lngRow, COL_医嘱内容)
                    If .Cell(flexcpData, lngRow - 1, COL_医嘱内容) <> "" Then
                        .TextMatrix(lngRow - 1, COL_医嘱内容) = .TextMatrix(lngRow - 1, COL_医嘱内容) & "(" & .Cell(flexcpData, lngRow - 1, COL_医嘱内容) & ")"
                    End If
                    '删除该行
                    .RemoveItem lngRow
                    blnRemove = True
                End If
            End If
                        
            If Not blnRemove Then
                '定位到先前行
                If Val(.TextMatrix(lngRow, COL_医嘱ID)) & "_" & .TextMatrix(lngRow, COL_NO) = strKey Then
                    lngPreRow = lngRow
                End If
                lngRow = lngRow + 1
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
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Function LoadBillList() As Boolean
'功能：读取当前可用的辅诊单据清单
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim vBill As TYPE_Bill
    
    On Error GoTo errH
    
    ReDim marrBill(0)
    
    '加载可用单据
    strSQL = "Select Distinct A.ID,A.编号,A.名称,A.说明" & _
        " From 病历文件目录 A,病历文件组成 B" & _
        " Where A.种类=5 And A.前提 IN([1],3)" & _
        " And A.ID=B.病历文件ID And B.填写时机 IN(1,2)" & _
        " Order by A.编号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint场合)
    If Not rsTmp.EOF Then
        ReDim marrBill(rsTmp.RecordCount)
        For i = 1 To rsTmp.RecordCount
            With vBill
                .ID = rsTmp!ID
                .名称 = rsTmp!名称
            End With
            marrBill(i) = vBill '第0的个不算
            rsTmp.MoveNext
        Next
    End If
    LoadBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
