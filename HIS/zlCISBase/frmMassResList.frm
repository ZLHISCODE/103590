VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmMassResList 
   Caption         =   "质控品管理"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   Icon            =   "frmMassResList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10650
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picRes 
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   4845
      ScaleHeight     =   3795
      ScaleWidth      =   6225
      TabIndex        =   4
      Top             =   360
      Width           =   6225
      Begin XtremeReportControl.ReportControl rptRes 
         Height          =   3405
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6030
         _Version        =   589884
         _ExtentX        =   10636
         _ExtentY        =   6006
         _StockProps     =   0
         BorderStyle     =   2
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
      Top             =   6060
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMassResList.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13705
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMassResList.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMassResList.frx":13B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMassResList.frx":1950
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
      Bindings        =   "frmMassResList.frx":1EEA
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMassResList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mColD
    图标 = 0: ID: 编码: 名称: 质控周期: 水平数: 使用部门
End Enum
Private Enum mColR
    图标 = 0: ID: 批号: 名称: 浓度: 方法: 开始日期: 结束日期
End Enum

Const conPane_Dev = 201
Const conPane_Res = 202
Const conPane_Edit = 203

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串
Private mfrmEdit As frmMassResEdit
Attribute mfrmEdit.VB_VarHelpID = -1

Private mintEditState As Integer    '当前编辑状态：0-非编辑状态,1-编辑状态
Private mlngDevId As Long, mlngResId As Long    '仪器id和物品id

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
    
    If gstrDBOwner = gstrDBUser Or InStr(mstrPrivs, "所有科室") > 0 Then
        '所有者，不限制
        gstrSql = "Select A.ID, A.编码, A.名称, D.名称 As 使用部门, Count(S.项目id) As 是否失控," & vbNewLine & _
                "       Decode(Nvl(A.质控周期, 0), 0, '', A.质控周期 || Nvl(A.周期单位, '天')) As 质控周期, A.质控水平数 As 水平数" & vbNewLine & _
                "From 检验仪器 A, 部门表 D, 检验仪器状态 S" & vbNewLine & _
                "Where A.使用小组id = D.ID(+) And A.ID = S.仪器id(+) And Nvl(A.微生物, 0) <> 1" & vbNewLine & _
                "Group By A.ID, A.编码, A.名称, D.名称, A.质控周期, A.周期单位, A.质控水平数"
    Else
        gstrSql = "Select A.ID, A.编码, A.名称, D.名称 As 使用部门, Count(S.项目id) As 是否失控," & vbNewLine & _
                "       Decode(Nvl(A.质控周期, 0), 0, '', A.质控周期 || Nvl(A.周期单位, '天')) As 质控周期, A.质控水平数 As 水平数" & vbNewLine & _
                "From 检验仪器 A," & vbNewLine & _
                "     (Select A.名称, A.ID" & vbNewLine & _
                "       From 部门表 A, 部门人员 B, 上机人员表 C" & vbNewLine & _
                "       Where A.ID = B.部门id And B.人员id = C.人员id And C.用户名 = User) D, 检验仪器状态 S" & vbNewLine & _
                "Where A.使用小组id = D.ID And A.ID = S.仪器id(+) And Nvl(A.微生物, 0) <> 1" & vbNewLine & _
                "Group By A.ID, A.编码, A.名称, D.名称, A.质控周期, A.周期单位, A.质控水平数"

    End If
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.rptDev.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptDev.Records.Add()
            If Val("" & !是否失控) = 0 Then
                rptRcd.AddItem("0").Icon = 0
            Else
                rptRcd.AddItem("0").Icon = 1
            End If
            rptRcd.AddItem CStr(!ID)
            rptRcd.AddItem CStr(!编码)
            rptRcd.AddItem CStr(!名称)
            rptRcd.AddItem CStr("" & !质控周期)
            rptRcd.AddItem CStr("" & !水平数)
            rptRcd.AddItem CStr("" & !使用部门)
            .MoveNext
        Loop
    End With
    Me.rptDev.Populate
    
    mlngDevId = 0
    If Me.rptDev.Rows.Count > 0 And (Me.rptDev.FocusedRow Is Nothing) Then
        Set Me.rptDev.FocusedRow = Me.rptDev.Rows(0)
        mlngDevId = Val(Me.rptDev.FocusedRow.Record(mColD.ID).Value)
    End If
    zlRefDev = Me.rptDev.Records.Count
    Me.stbThis.Panels(2).Text = "共有" & Me.rptDev.Records.Count & "台仪器"
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefDev = Me.rptDev.Records.Count
End Function

Private Function zlRefRes(Optional lngResID As Long) As Long
    '功能：刷新装入当前仪器的质控品
    Dim rsTemp As New ADODB.Recordset
    
    gstrSql = "Select ID, 名称, 批号, Decode(非定值, 1, '1-非定值', '0-定值') As 类型," & vbNewLine & _
            "       浓度 || Decode(Nvl(水平, 0), 0, '', ' (水平' || 水平 || ')') As 浓度, 方法, 开始日期, 结束日期" & vbNewLine & _
            "From 检验质控品" & vbNewLine & _
            "Where 仪器id = [1]"
    
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevId)
    Err = 0: On Error GoTo 0
    Me.rptRes.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptRes.Records.Add()
            rptRcd.AddItem("1").Icon = 2
            rptRcd.AddItem CStr(!ID)
            rptRcd.AddItem CStr(!批号)
            rptRcd.AddItem CStr(!名称)
            rptRcd.AddItem CStr("" & !浓度)
            rptRcd.AddItem CStr("" & !方法)
            rptRcd.AddItem CStr(Format(!开始日期, "yyyy-MM-dd; ; ;"))
            rptRcd.AddItem CStr(Format(!结束日期, "yyyy-MM-dd; ; ;"))
            .MoveNext
        Loop
    End With
    Me.rptRes.Populate
    
    If lngResID <> 0 Then
        For Each rptRow In Me.rptRes.Rows
            If Val(rptRow.Record(mColR.ID).Value) = lngResID Then
                Set Me.rptRes.FocusedRow = rptRow
                mlngResId = Val(Me.rptRes.FocusedRow.Record(mColR.ID).Value)
                Exit For
            End If
        Next
    End If
    If Me.rptRes.FocusedRow Is Nothing Then
        If Me.rptRes.Rows.Count > 0 Then
            Set Me.rptRes.FocusedRow = Me.rptRes.Rows(0)
            mlngResId = Val(Me.rptRes.FocusedRow.Record(mColR.ID).Value)
        Else
            mlngResId = 0: Call rptRes_SelectionChanged
        End If
    End If
    zlRefRes = Me.rptRes.Records.Count
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefRes = Me.rptRes.Records.Count
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    If Me.rptRes.Records.Count = 0 Then Exit Sub
    '-------------------------------------------------
    '复制数据表格
    If zlControl.RPTCopyToVSF(Me.rptRes, Me.vfgList) Is Nothing Then Exit Sub
    
    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "检验质控品清单"
    
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
            mlngResId = lngRetuId: Call zlRefRes(mlngResId)
            mintEditState = 0: Me.picDev.Enabled = True: Me.picRes.Enabled = True: Me.rptRes.SetFocus
        End If
        
    Case conMenu_Edit_Untread:
        mfrmEdit.zlEditCancel: Call zlRefRes(mlngResId)
        mintEditState = 0: Me.picDev.Enabled = True: Me.picRes.Enabled = True: Me.rptRes.SetFocus

    Case conMenu_Edit_NewItem
        If mlngDevId = 0 Then Exit Sub
        If mfrmEdit.zlEditStart(True, mlngResId, mlngDevId) Then
            mintEditState = 1: Me.picDev.Enabled = False: Me.picRes.Enabled = False
        End If
        Me.dkpMan.FindPane(conPane_Edit).Select
    Case conMenu_Edit_Modify
        If mlngDevId = 0 Then Exit Sub
        If mlngResId = 0 Then Exit Sub
        If mfrmEdit.zlEditStart(False, mlngResId, mlngDevId) Then
            mintEditState = 1: Me.picDev.Enabled = False: Me.picRes.Enabled = False
        End If
        Me.dkpMan.FindPane(conPane_Edit).Select
    Case conMenu_Edit_Delete
        Dim strMsg As String
        If mlngResId = 0 Then Exit Sub
        With Me.rptRes
            strMsg = "真的删除该检验质控品吗？" & vbCrLf & "——"
            strMsg = strMsg & .FocusedRow.Record(mColR.名称).Value & "  批号:" & .FocusedRow.Record(mColR.批号).Value
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSql = "Zl_检验质控品_Edit(3," & mlngResId & ")"
                Err = 0: On Error GoTo ErrHand
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
                
                Err = 0: On Error GoTo 0
                mlngResId = 0: lngRetuId = .FocusedRow.Index
                If .Rows.Count > lngRetuId + 1 Then
                    mlngResId = .Rows(lngRetuId + 1).Record(mColR.ID).Value
                ElseIf lngRetuId > 0 Then
                    mlngResId = .Rows(lngRetuId - 1).Record(mColR.ID).Value
                End If
                Call zlRefRes(mlngResId)
            End If
        End With
        Exit Sub
    Case conMenu_Edit_Send
        frmMassResCopy.Show vbModal, Me
        Call zlRefRes(mlngResId)
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
        Call zlRefRes(mlngResId)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
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
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptRes.Records.Count > 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (mintEditState <> 0)
    Case conMenu_Edit_NewItem
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0) And mlngDevId <> 0
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0 And mlngResId <> 0)
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
        If mfrmEdit Is Nothing Then Set mfrmEdit = New frmMassResEdit
        Item.Handle = mfrmEdit.hWnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    
    mintEditState = 0
    mlngDevId = 0: mlngResId = 0
    
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
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "复制(&P)"): cbrControl.BeginGroup = True
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "复制"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panThis As Pane, panSub1 As Pane
    
    If mfrmEdit Is Nothing Then Set mfrmEdit = New frmMassResEdit
    
    Set panThis = dkpMan.CreatePane(conPane_Dev, 240, 600, DockLeftOf, Nothing)
    panThis.Title = "检验仪器列表"
    panThis.Options = PaneNoCaption
    Set panThis = dkpMan.CreatePane(conPane_Res, 600, 300, DockRightOf, Nothing)
    panThis.Title = "质控品列表"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panSub1 = dkpMan.CreatePane(conPane_Edit, 600, 400, DockBottomOf, panThis)
    panSub1.Title = "质控品信息"
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
        Set rptCol = .Columns.Add(mColD.编码, "编码", 50, False): rptCol.Editable = False: rptCol.Groupable = False
        .SortOrder.Add rptCol: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColD.名称, "名称", 130, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColD.质控周期, "质控周期", 55, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColD.水平数, "水平数", 45, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColD.使用部门, "使用部门", 80, True): rptCol.Editable = False: rptCol.Groupable = False
        
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
    End With
    
    '-----------------------------------------------------
    '物品列表的设置
    With Me.rptRes
        .AutoColumnSizing = True
        Set rptCol = .Columns.Add(mColR.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColR.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mColR.批号, "批号", 60, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mColR.名称, "名称", 150, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.浓度, "浓度", 90, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.方法, "方法", 90, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.开始日期, "开始日期", 70, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.结束日期, "结束日期", 70, False): rptCol.Editable = False: rptCol.Groupable = False
        
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
    End With
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '数据装入
    Call zlRefDev

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
    With Me.rptRes
        .Left = Me.picRes.ScaleLeft: .Width = Me.picRes.ScaleWidth - .Left
        .Top = Me.picRes.ScaleTop: .Height = Me.picRes.ScaleHeight - .Top
    End With
End Sub

Private Sub rptDev_SelectionChanged()
    If Me.rptDev.FocusedRow Is Nothing Then
        mlngDevId = 0
    Else
        mlngDevId = Me.rptDev.FocusedRow.Record.Item(mColD.ID).Value
    End If
    Call zlRefRes
End Sub

Private Sub rptRes_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptRes.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptRes.FocusedRow Is Nothing Then Exit Sub
    If Me.rptRes.FocusedRow.GroupRow Then Exit Sub
    Call rptRes_RowDblClick(Me.rptRes.FocusedRow, Me.rptRes.FocusedRow.Record.Item(mColR.ID))
End Sub

Private Sub rptRes_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
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

Private Sub rptRes_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub rptRes_SelectionChanged()
'    If Me.Visible = False Then Exit Sub
    If Me.rptRes.FocusedRow Is Nothing Then
        mlngResId = 0
    Else
        mlngResId = Me.rptRes.FocusedRow.Record.Item(mColD.ID).Value
    End If
    Call mfrmEdit.zlRefresh(mlngResId, mlngDevId)
End Sub
