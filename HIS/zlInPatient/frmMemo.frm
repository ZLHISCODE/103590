VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMemo 
   Caption         =   "病人备注编辑"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8295
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picUserInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   8295
      TabIndex        =   2
      Top             =   0
      Width           =   8295
      Begin VB.Label lblUserInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   720
         Left            =   120
         Picture         =   "frmMemo.frx":6852
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5430
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   635
      SimpleText      =   $"frmMemo.frx":851C
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMemo.frx":8563
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9551
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
   Begin VSFlex8Ctl.VSFlexGrid vsfMemo 
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
      _cx             =   3625
      _cy             =   3625
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   2640
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":8DF7
            Key             =   "紧急标志"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":9191
            Key             =   "当前"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":F9F3
            Key             =   "指示器"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":16255
            Key             =   "附件"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":167EF
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":16B89
            Key             =   "标志"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":16F23
            Key             =   "单据"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":172BD
            Key             =   "紧急"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":17657
            Key             =   "图标"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":179F1
            Key             =   "新申请"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":1E253
            Key             =   "待手术"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":24AB5
            Key             =   "手术中"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":2B317
            Key             =   "已手术"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":31B79
            Key             =   "已拒绝"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":383DB
            Key             =   "复苏中"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":3EC3D
            Key             =   "已执行"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":4549F
            Key             =   "待复苏"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":4BD01
            Key             =   "已接收"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":52563
            Key             =   "分类"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":528FD
            Key             =   "分类_选中"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":52C97
            Key             =   "分类_固定"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":594F9
            Key             =   "项目"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":5FD5B
            Key             =   "审核"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":665BD
            Key             =   "方案"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":6CE1F
            Key             =   "收费"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":6D831
            Key             =   "事件"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   1440
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPrivs As String
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mblnAllowEdit As Boolean
Private mblnDataChanged As Boolean
Private mbln结清 As Boolean
Private mrsPatiInfo As ADODB.Recordset

Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Public Event AfterDataChanged()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, objControl As CommandBarControl
    
    On Err GoTo errHandle:
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hWnd)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '退出
        Unload Me
        
    Case conMenu_Edit_Modify  '编辑
        mclsVsf.AllowEdit = True
        Call vsfMemo.Select(vsfMemo.Row, vsfMemo.ColIndex("备注信息"))
    Case conMenu_Edit_Save '保存
        If Not SaveData Then Exit Sub
        mclsVsf.AllowEdit = False
        DataChanged = False
        Call LoadData
    Case conMenu_Edit_Untread  '取消
        If MsgBox("您已经对该病人备注信息做了修改，是否保存？", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            DataChanged = False
            mclsVsf.AllowEdit = False
            Call LoadData
        Else
            If Not SaveData Then Exit Sub
            mclsVsf.AllowEdit = False
            DataChanged = False
            Call LoadData
        End If
    Case conMenu_Edit_Delete '删除
        If mblnAllowEdit = False Then Exit Sub
        If Not CheckWOver(vsfMemo.Row) Then
            MsgBox "您不能删除非本人完成的项目！", vbInformation, gstrSysName
            Exit Sub
        End If
        If MsgBox("您确定要删除所选备注信息吗？", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        mclsVsf.AllowEdit = True
        gstrSQL = "ZL_病人备注信息_Delete(" & Val(vsfMemo.TextMatrix(vsfMemo.Row, vsfMemo.ColIndex("ID"))) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Call mclsVsf.DeleteRow(vsfMemo.Row)
        mclsVsf.AllowEdit = False
    Case conMenu_View_Refresh
        DataChanged = False
        mclsVsf.AllowEdit = False
        Call LoadData
    End Select
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function SaveData() As Boolean
    '保存数据
    Dim i As Integer, iRow As Integer
    Dim blnTrans As Boolean, intTmp As Integer
    Dim strSQL() As String, str编码 As String, str名称 As String
    Dim lngColor As Long '-214748363
    On Error GoTo errHandle
    
    With vsfMemo
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("备注信息"))) <> "" Then
                If zlCommFun.ActualLen(Trim(.TextMatrix(i, .ColIndex("备注信息")))) > 200 Then
                    MsgBox "备注信息过长，最多允许 100 个汉字或 200 个字符。", vbInformation, gstrSysName
                    vsfMemo.Row = i
                    vsfMemo.SetFocus: Exit Function
                End If
                If .TextMatrix(i, .ColIndex("更改标志")) = "1" And Val(.TextMatrix(i, .ColIndex("ID"))) > 0 Then
                    ReDim Preserve strSQL(intTmp)
                    strSQL(UBound(strSQL)) = "Zl_病人备注信息_Update(" & Val(.TextMatrix(i, .ColIndex("ID"))) & "," & mlng病人ID & "," & mlng主页ID & ",'" & _
                            Trim(.TextMatrix(i, .ColIndex("备注信息"))) & "',to_Date('" & _
                            Trim(.TextMatrix(i, .ColIndex("登记时间"))) & "','YYYY-MM-DD HH24:MI:SS'), '" & _
                            Trim(.TextMatrix(i, .ColIndex("登记人"))) & "',1," & Val(.TextMatrix(i, .ColIndex("是否完成"))) & ",to_Date('" & _
                            Trim(.TextMatrix(i, .ColIndex("完成时间"))) & "','YYYY-MM-DD HH24:MI:SS'), '" & _
                            Trim(.TextMatrix(i, .ColIndex("完成人"))) & "')"
                    intTmp = intTmp + 1
                ElseIf .TextMatrix(i, .ColIndex("更改标志")) = "1" And Val(.TextMatrix(i, .ColIndex("ID"))) = 0 Then
                    ReDim Preserve strSQL(intTmp)
                    strSQL(UBound(strSQL)) = "Zl_病人备注信息_Update(" & zlDatabase.GetNextId("病人备注信息") & "," & mlng病人ID & "," & mlng主页ID & ",'" & _
                            Trim(.TextMatrix(i, .ColIndex("备注信息"))) & "',to_Date('" & _
                            Trim(.TextMatrix(i, .ColIndex("登记时间"))) & "','YYYY-MM-DD HH24:MI:SS'), '" & _
                            Trim(.TextMatrix(i, .ColIndex("登记人"))) & "',0," & Val(.TextMatrix(i, .ColIndex("是否完成"))) & ",to_Date('" & _
                            Trim(.TextMatrix(i, .ColIndex("完成时间"))) & "','YYYY-MM-DD HH24:MI:SS'), '" & _
                            Trim(.TextMatrix(i, .ColIndex("完成人"))) & "')"
                    intTmp = intTmp + 1
                End If
                .TextMatrix(i, .ColIndex("更改标志")) = "0"
            End If
        Next
    End With
    
    If intTmp > 0 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(strSQL) To UBound(strSQL)
            Call zlDatabase.ExecuteProcedure(strSQL(i), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    SaveData = True
    
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    With Me.picUserInfo
         .Left = lngLeft + 30: .Top = lngTop + 30
         .width = lngRight - lngLeft - 30
         '.Height = lngBottom - .Top - 30 ' - IIf(mblnAllowEdit, stbThis.Height, 0)
    End With
    With Me.vsfMemo
'        If mblnAllowEdit = True Then
'            .Left = lngLeft + 30: .Top = lngTop + 30
'            .Width = lngRight - lngLeft - 30
'            .Height = lngBottom - .Top - 30 ' - IIf(mblnAllowEdit, stbThis.Height, 0)
'        Else
            .Left = lngLeft + 30: .Top = picUserInfo.Height + lngTop + 30
            .width = lngRight - lngLeft - 30
            .Height = lngBottom - .Top - 30 ' - IIf(mblnAllowEdit, stbThis.Height, 0)
'        End If
    End With
    
    mclsVsf.AppendRows = True
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
        
    '-------------------------------
    Case conMenu_Edit_Save '保存
        Control.Enabled = DataChanged
    Case conMenu_Edit_Untread '取消
        Control.Enabled = DataChanged
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    mblnAllowEdit = True
    mbln结清 = True

    '菜单工具栏
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call InitCommandBar
    Call InitVSFlexGrid
    Call LoadData
    '没有编辑权限是不允许编辑
    If InStr(mstrPrivs, "病人备注编辑") = 0 Then mblnAllowEdit = False
    '出院病人不允许编辑
    If Not IsNull(mrsPatiInfo!出院日期) And mbln结清 Then mblnAllowEdit = False
    Me.lblUserInfo = "姓名：" & mrsPatiInfo!姓名 & "     " & "性别：" & mrsPatiInfo!性别 & "   " & "年龄：" & mrsPatiInfo!年龄 & "   " & "住院号：" & mrsPatiInfo!住院号
    If Not mblnAllowEdit Then
        Me.Caption = "病人备注信息(当前用户：" & UserInfo.姓名 & ")"
        'Me.picUserInfo.Visible = True
        'Me.lblUserInfo = "姓名：" & mrsPatiInfo!姓名 & "     " & "性别：" & mrsPatiInfo!性别 & "   " & "年龄：" & mrsPatiInfo!年龄 & "   " & "住院号：" & mrsPatiInfo!住院号
        Me.cbsMain.ActiveMenuBar.Visible = False
        For i = 2 To cbsMain.Count
            cbsMain(i).Visible = False
        Next
        stbThis.Visible = False
        Me.cbsMain.RecalcLayout
    End If
End Sub

Private Sub InitVSFlexGrid()
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsfMemo, True, True, ils16)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[序号]", False)
        Call .AppendColumn("更改标志", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTLong, , , True, , , True)
        Call .AppendColumn("病人ID", 0, flexAlignLeftCenter, flexDTLong, , , True, , , True)
        Call .AppendColumn("主页ID", 0, flexAlignLeftCenter, flexDTLong, , , True, , , True)
        
        Call .AppendColumn("备注信息", 5000, flexAlignLeftCenter, flexDTString, , "内容", True, , , , True)
        Call .AppendColumn("登记时间", 2000, flexAlignLeftCenter, flexDTString, , , True)
        Call .AppendColumn("登记人", 800, flexAlignLeftCenter, flexDTString, , , True)
        '51338,刘鹏飞,2012-09-04,添加是否完成、完成时间、完成人
        Call .AppendColumn("是否完成", 1000, flexAlignCenterCenter, flexDTBoolean, , "是否完成", False, True)
        Call .AppendColumn("完成时间", 2000, flexAlignLeftCenter, flexDTString, , , True)
        Call .AppendColumn("完成人", 800, flexAlignLeftCenter, flexDTString, , , True)
            
        If InStr(mstrPrivs, "病人备注编辑") And mblnAllowEdit Then
            Call .InitializeEdit(True, True, True)
            Call .InitializeEditColumn(.ColIndex("备注信息"), True, vbVsfEditText, , 200)
            Call .InitializeEditColumn(.ColIndex("是否完成"), True, vbVsfEditCheck, , 1)
        End If
        .SysHidden(.ColIndex("更改标志")) = True
        .IndicatorMode = 2
        .AppendRows = True
        .AllowEdit = False
    End With
End Sub

Private Sub LoadData()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    Dim lngColor As Long
    
    On Error GoTo errHandle
    ' 获取病人是否结清
    strSQL = " Select Nvl(sum(费用余额),0) 费用余额 From 病人余额 Where 病人ID=[1] And 性质=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    If rsTmp.RecordCount > 0 Then
        mbln结清 = Not CBool(Val("" & rsTmp!费用余额))
    End If
    
    mclsVsf.ClearGrid
            
    zlCommFun.ShowFlash "正在装载数据，请稍等..."
    
    LockWindowUpdate vsfMemo.hWnd
    Set mrsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID)
    
    '51338,刘鹏飞,2012-09-04,添加是否完成、完成时间、完成人
    strSQL = "Select id, 病人id, 主页id, 内容, 登记时间, 登记人,是否完成,完成时间,完成人 From 病人备注信息 Where 病人id = [1] And 主页id = [2] Order By Id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.BOF = False Then
        Call mclsVsf.LoadGrid(rsTmp)
    End If
    For lngColor = vsfMemo.FixedRows To vsfMemo.Rows - 1
        If IsDate(vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("登记时间"))) = True Then vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("登记时间")) = Format(vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("登记时间")), "YYYY-MM-DD HH:mm:ss")
        If IsDate(vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("完成时间"))) = True Then vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("完成时间")) = Format(vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("完成时间")), "YYYY-MM-DD HH:mm:ss")
    Next
    LockWindowUpdate 0
    
    zlCommFun.StopFlash
    DataChanged = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub InitCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = zlCommFun.GetPubIcons

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False) '固有
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")

        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "编辑(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)") '固有
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)") '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False) '固有
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)") '固有
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName) '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True '固有
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有

        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "编辑"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")

        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出") '固有
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub SetOver(ByVal lngRow As Long)
    '51338,刘鹏飞,2012-09-04,添加是否完成、完成时间、完成人
    If vsfMemo.TextMatrix(lngRow, vsfMemo.ColIndex("是否完成")) = "-1" Then
        vsfMemo.TextMatrix(lngRow, vsfMemo.ColIndex("完成时间")) = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
        vsfMemo.TextMatrix(lngRow, vsfMemo.ColIndex("完成人")) = UserInfo.姓名
    Else
        vsfMemo.TextMatrix(lngRow, vsfMemo.ColIndex("完成时间")) = ""
        vsfMemo.TextMatrix(lngRow, vsfMemo.ColIndex("完成人")) = ""
    End If
End Sub

Private Sub mclsVSF_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub mclsVSF_AfterNewRow(ByVal Row As Long, Col As Long)
    If mclsVsf.AllowEdit = False Then Exit Sub
    vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("更改标志")) = "1"
    vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("登记时间")) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("登记人")) = UserInfo.姓名
    Call SetOver(Row)
End Sub

Private Sub mclsVSF_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo errHandle
    If Not CheckWOver(vsfMemo.Row) Then
        MsgBox "您不能删除非本人完成的项目！", vbInformation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    gstrSQL = "ZL_病人备注信息_Delete(" & Val(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("ID"))) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mclsVSF_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = Trim(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("备注信息"))) = ""
End Sub

Private Sub vsfMemo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '编辑处理
    Call mclsVsf.AfterEdit(Row, Col)
    vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("更改标志")) = "1"
    If Col = vsfMemo.ColIndex("是否完成") Then Call SetOver(Row)
    DataChanged = True
End Sub

Private Sub vsfMemo_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Call mclsVsf.AfterMoveColumn(Col, Position)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsfMemo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsfMemo_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsfMemo_AfterSort(ByVal Col As Long, Order As Integer)
    Call mclsVsf.RestoreRow(mclsVsf.SaveKey)
    vsfMemo.ShowCell vsfMemo.Row, vsfMemo.Col
End Sub

Private Sub vsfMemo_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsfMemo_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow = NewRow Then Exit Sub
    vsfMemo.ForeColorSel = vsfMemo.Cell(flexcpForeColor, NewRow, 0)
End Sub

Private Sub vsfMemo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsfMemo_ChangeEdit()
    With vsfMemo
        Select Case .Col
        Case .ColIndex("备注信息")
            .TextMatrix(.Row, .Col) = .EditText
            .TextMatrix(.Row, .ColIndex("更改标志")) = "1"
            .TextMatrix(.Row, .ColIndex("登记时间")) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(.Row, .ColIndex("登记人")) = UserInfo.姓名
            Call SetOver(.Row)
            DataChanged = True
        End Select
    End With
End Sub

Private Sub vsfMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsfMemo_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytRet As Byte
    Dim strDoctor As String
    Dim bln麻醉人员 As Boolean
    
    With vsfMemo
        If KeyCode = vbKeyReturn Then
            
            If InStr(.EditText, "'") > 0 Then
                KeyCode = 0
                .EditText = ""
                Exit Sub
            End If
            
            strText = Trim(.EditText)
                                
            Select Case Col
            Case .ColIndex("备注信息")
                    DataChanged = True
                    .TextMatrix(Row, .ColIndex("更改标志")) = "1"
                    .TextMatrix(Row, .ColIndex("登记时间")) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                    .TextMatrix(Row, .ColIndex("登记人")) = UserInfo.姓名
                    Call SetOver(Row)
'                End If
            Case Else
                Call mclsVsf.LocationNextCell
            End Select
            
            mclsVsf.LocationNextCell
            
            mclsVsf.SetFocus , , True
        End If
    End With
End Sub

Private Sub vsfMemo_KeyPress(KeyAscii As Integer)
    '编辑处理
    If InStr("'", Chr(KeyAscii)) > 0 Then Exit Sub
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsfMemo_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsfMemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsfMemo.MouseRow, vsfMemo.MouseCol)
    End Select
End Sub

Private Sub vsfMemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call gclsBase.SendLMouseButton(vsfMemo.hWnd, X, Y)
        If mclsVsf.MoveColumn = False Then
            RaiseEvent MouseUp(Button, Shift, X, Y)
        End If
    End Select
End Sub

Private Sub vsfMemo_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsfMemo_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '51338,刘鹏飞,2012-09-04,添加是否完成、完成时间、完成人
    '已经标记完成的项目只有本人才能进行修改
    If Not CheckWOver(Row) And mclsVsf.AllowEdit = True Then Cancel = True: Exit Sub
    '编辑处理
    If Col = vsfMemo.ColIndex("备注信息") Then vsfMemo.EditMaxLength = 200
    If Col = vsfMemo.ColIndex("是否完成") And Trim(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("备注信息"))) = "" Then Cancel = True: Exit Sub
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsfMemo_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
    If Cancel Then Exit Sub
End Sub

Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    
    ShowMe = gblnOK
End Function

Private Function CheckWOver(ByVal Row As Long)
'对于已经完成的项目检查是否是本人完成
    If Val(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("是否完成"))) = -1 Then '标明已经完成
        If Trim(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("完成人"))) = "" Then Trim(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("完成人"))) = UserInfo.姓名
        If Trim(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("完成人"))) <> Trim(UserInfo.姓名) Then Exit Function
    End If
    CheckWOver = True
End Function
