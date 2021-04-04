VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMicrobeAntiRef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "细菌抗生素参考"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11865
   Icon            =   "frmMicrobeAntiRef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   11865
   StartUpPosition =   2  '屏幕中心
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   5685
      Left            =   165
      TabIndex        =   11
      Top             =   135
      Width           =   3030
      _Version        =   589884
      _ExtentX        =   5345
      _ExtentY        =   10028
      _StockProps     =   0
      BorderStyle     =   1
   End
   Begin VB.PictureBox picVfg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3270
      Left            =   3255
      ScaleHeight     =   3240
      ScaleWidth      =   8085
      TabIndex        =   14
      Top             =   405
      Width           =   8115
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   3795
         Left            =   45
         TabIndex        =   15
         Top             =   45
         Width           =   8460
         _cx             =   14922
         _cy             =   6694
         Appearance      =   0
         BorderStyle     =   0
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
         Cols            =   6
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
   End
   Begin VB.PictureBox picEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   3330
      ScaleHeight     =   1905
      ScaleWidth      =   8430
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3735
      Width           =   8460
      Begin VB.ComboBox cbo结果 
         Height          =   300
         Index           =   2
         Left            =   6270
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   900
         Width           =   1215
      End
      Begin VB.ComboBox cbo结果 
         Height          =   300
         Index           =   1
         Left            =   3600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   900
         Width           =   1215
      End
      Begin VB.ComboBox cbo结果 
         Height          =   300
         Index           =   0
         Left            =   930
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   900
         Width           =   1215
      End
      Begin VB.CommandButton cmd抗生素 
         Caption         =   "…"
         Height          =   300
         Left            =   7845
         TabIndex        =   13
         Top             =   180
         Width           =   300
      End
      Begin VB.TextBox txt抗生素 
         Height          =   300
         Left            =   930
         TabIndex        =   1
         Top             =   195
         Width           =   6900
      End
      Begin VB.ComboBox cbo判断方式 
         Height          =   300
         Left            =   6270
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   540
         Width           =   1890
      End
      Begin VB.ComboBox cbo方法 
         Height          =   300
         ItemData        =   "frmMicrobeAntiRef.frx":000C
         Left            =   930
         List            =   "frmMicrobeAntiRef.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox txt参考值 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   2760
         MaxLength       =   13
         TabIndex        =   3
         Top             =   540
         Width           =   900
      End
      Begin VB.TextBox txt参考值 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   3900
         MaxLength       =   13
         TabIndex        =   4
         Top             =   540
         Width           =   900
      End
      Begin VB.TextBox txt备注 
         Height          =   300
         Left            =   930
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1245
         Width           =   7230
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "高于参考"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   5070
         TabIndex        =   21
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "在参考范围内"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   2325
         TabIndex        =   19
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "低于参考"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   17
         Top             =   960
         Width           =   720
      End
      Begin XtremeCommandBars.CommandBars cbsThis 
         Left            =   180
         Top             =   1410
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
      Begin VB.Label lbl抗生素 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "抗生素"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   12
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl方式 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "参考判断方式"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5070
         TabIndex        =   10
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label lbl方法 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药敏方法"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl参考 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "参考           ～"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2340
         TabIndex        =   8
         Top             =   600
         Width           =   1530
      End
      Begin VB.Label lbl备注 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   180
         Left            =   105
         TabIndex        =   7
         Top             =   1305
         Width           =   360
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMicrobeAntiRef.frx":0010
      Left            =   2745
      Top             =   45
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMicrobeAntiRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng细菌id As Long
Private mlngItemID As Long '上次选择的分组

Private Enum mCol
    图标 = 0: 分组Id: 编号: 名称: 英文
    ID = 0: 编码: 中文名: 英文名: 药敏方法: 参考低值: 参考高值: 参考: 判断方式: 备注: 关键字: 低值结果: 中间结果: 高值结果
End Enum

Private Const Dkp_ID_Rpt As Integer = 1
Private Const Dkp_ID_vfg As Integer = 2
Private Const Dkp_ID_Edit As Integer = 3
Private cbrControl As CommandBarControl
Private mblnEdit As Boolean

Private Sub cbo方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbo判断方式_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCurRow As Long, lngRow As Long, lngCol As Long
    Dim str抗生素 As String, rsTmp As ADODB.Recordset, strSQL As String
    Dim str关键字 As String
    
    On Error GoTo errHandle
    With Me.vfgList
        Select Case Control.ID
        Case conMenu_Edit_NewItem
            .Rows = .Rows + 1: .Row = .Rows - 1
        Case conMenu_Edit_Delete
            str关键字 = .TextMatrix(.Row, mCol.关键字)
            If str关键字 <> "" Then
                strSQL = "Zl_检验细菌抗生素参考_Edit(2," & mlng细菌id & "," & mlngItemID & "," & Split(str关键字, ",")(0) & "," & Split(str关键字, ",")(1) & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
            If .Row = .Rows - 1 Then
                .Rows = .Rows - 1: .Row = .Rows - 1
            Else
                lngCurRow = .Row
                For lngRow = lngCurRow To .Rows - 2
                    For lngCol = 0 To .Cols - 1
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow + 1, lngCol)
                    Next
                Next
                .Rows = .Rows - 1
            End If
        Case conMenu_Edit_Adjust
            If Val(txt抗生素.Tag) <> 0 Then
                For lngRow = .FixedRows To .Rows - 1
                    If lngRow <> .Row Then
                        If .TextMatrix(lngRow, mCol.ID) = Val(txt抗生素.Tag) And .TextMatrix(lngRow, mCol.药敏方法) = Me.cbo方法.Text Then
                            MsgBox "已有相同记录存在,不能更新!", vbQuestion, Me.Caption
                            Exit Sub
                        End If
                    End If
                Next
            End If
            
            If Me.cbo方法.Text = "" Then
                MsgBox "必须设置一个药敏方法后，才能更新参数！", vbInformation, Me.Caption
                Me.cbo方法.SetFocus
            End If
            
            .TextMatrix(.Row, mCol.ID) = Val(txt抗生素.Tag)
            .TextMatrix(.Row, mCol.编码) = ""
            .TextMatrix(.Row, mCol.中文名) = ""
            .TextMatrix(.Row, mCol.英文名) = ""
            
            If Val(txt抗生素.Tag) <> 0 Then
                strSQL = "Select B.编码, B.中文名, B.英文名 From 检验用抗生素 B Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(txt抗生素.Tag))
                Do Until rsTmp.EOF
                    .TextMatrix(.Row, mCol.编码) = "" & rsTmp!编码
                    .TextMatrix(.Row, mCol.中文名) = "" & rsTmp!中文名
                    .TextMatrix(.Row, mCol.英文名) = "" & rsTmp!英文名
                    rsTmp.MoveNext
                Loop
            End If
            
            If Me.cbo方法.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.药敏方法) = ""
            Else
                .TextMatrix(.Row, mCol.药敏方法) = Me.cbo方法.Text
                .TextMatrix(.Row, mCol.关键字) = .TextMatrix(.Row, mCol.ID) & "," & Me.cbo方法.ListIndex + 1
            End If
            
            .TextMatrix(.Row, mCol.参考低值) = IIf(IsNumeric(txt参考值(0)), Me.txt参考值(0), "")
            .TextMatrix(.Row, mCol.参考高值) = IIf(IsNumeric(txt参考值(1)), Me.txt参考值(1), "")
            
            If .TextMatrix(.Row, mCol.参考低值) = "" Or "" & .TextMatrix(.Row, mCol.参考高值) = "" Then
                .TextMatrix(.Row, mCol.参考) = FormatDecimal(.TextMatrix(.Row, mCol.参考低值)) & FormatDecimal(.TextMatrix(.Row, mCol.参考高值))
            Else
                .TextMatrix(.Row, mCol.参考) = FormatDecimal(.TextMatrix(.Row, mCol.参考低值)) & "～" & FormatDecimal(.TextMatrix(.Row, mCol.参考高值))
            End If

            If Me.cbo判断方式.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.判断方式) = ""
            Else
                .TextMatrix(.Row, mCol.判断方式) = Me.cbo判断方式.Text
            End If
            .TextMatrix(.Row, mCol.备注) = DelInvalidChar(Trim(Me.txt备注.Text), "'")
            .TextMatrix(.Row, mCol.低值结果) = cbo结果(0).Text
            .TextMatrix(.Row, mCol.中间结果) = cbo结果(1).Text
            .TextMatrix(.Row, mCol.高值结果) = cbo结果(2).Text
            mblnEdit = True
        Case conMenu_Edit_Save
            Call zlSaveData
            Call initVfg(mlngItemID)
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_Delete, conMenu_Edit_Adjust: Control.Enabled = Me.vfgList.Row >= Me.vfgList.FixedRows
    Case conMenu_Edit_Save: Control.Enabled = mblnEdit
    End Select

End Sub

Private Sub zlSaveData()
    Dim lngRow As Long, strSQL As String, strDelSQL As String, str关键字 As String
    Dim str低 As String, str中 As String, str高 As String
    Dim strValue As String
    
    On Error GoTo errHandle
    With vfgList
        For lngRow = .FixedCols To .Rows - 1
            If Val(.TextMatrix(lngRow, mCol.ID)) <> 0 Then
                strSQL = "Zl_检验细菌抗生素参考_Edit(1," & mlng细菌id & "," & mlngItemID & "," & Val(.TextMatrix(lngRow, mCol.ID)) & "," & _
                         Get药敏方法(.TextMatrix(lngRow, mCol.药敏方法)) & ","
                If IsNumeric(.TextMatrix(lngRow, mCol.参考低值)) = True Then
                    
                    strValue = .TextMatrix(lngRow, mCol.参考低值)
                    If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                        MsgBox "第" & lngRow & "参考值太大或精度太高！", vbInformation, gstrSysName
                        Me.txt抗生素.SetFocus: Exit Sub
                    End If
                    strSQL = strSQL & strValue & ","
                Else
                    strSQL = strSQL & "Null,"
                End If
                
                If IsNumeric(.TextMatrix(lngRow, mCol.参考高值)) = True Then
                    
                    strValue = .TextMatrix(lngRow, mCol.参考高值)
                    If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                        MsgBox "第" & lngRow & "参考值太大或精度太高！", vbInformation, gstrSysName
                        Me.txt抗生素.SetFocus:  Exit Sub
                    End If
                    strSQL = strSQL & strValue & ","
                Else
                    strSQL = strSQL & "Null,"
                End If
                strSQL = strSQL & Get判断方式(.TextMatrix(lngRow, mCol.判断方式)) & ",'" & .TextMatrix(lngRow, mCol.备注) & "'"
                str关键字 = .TextMatrix(lngRow, mCol.关键字)
                str低 = .TextMatrix(lngRow, mCol.低值结果)
                str中 = .TextMatrix(lngRow, mCol.中间结果)
                str高 = .TextMatrix(lngRow, mCol.高值结果)
                strSQL = strSQL & ",'" & str低 & "','" & str中 & "','" & str高 & "')"
                
                If Right(Trim(str关键字), 1) <> "," Then
                    If str关键字 <> "" Then
                        strDelSQL = "Zl_检验细菌抗生素参考_Edit(2," & mlng细菌id & "," & mlngItemID & "," & Split(str关键字, ",")(0) & "," & Split(str关键字, ",")(1) & ")"
                        Call zlDatabase.ExecuteProcedure(strDelSQL, Me.Caption)
                    End If
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                End If
            End If
        Next
    End With
    mblnEdit = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cmd抗生素_Click()
    Dim rsTemp As ADODB.Recordset
    Dim blnReturn As Boolean
    On Error GoTo errHandle
    gstrSql = "Select B.ID, B.编码, B.中文名, B.英文名, 药敏方法" & vbNewLine & _
            "From 检验用抗生素 B, 检验抗生素用药 A" & vbNewLine & _
            "Where A.抗生素id = B.Id  And A.抗生素分组id = [1]"
    'Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "抗生素选择", True, "", "请选择抗生素", False, False, False, 0, 0, 0, blnReturn, True, False, mlngItemID)
    If blnReturn = False Then
        If rsTemp.RecordCount > 0 Then
            txt抗生素.Tag = rsTemp!ID
            txt抗生素.Text = "(" & rsTemp!编码 & ")" & rsTemp!中文名
            lbl抗生素.Tag = txt抗生素.Text '用于恢复显示
        Else
            txt抗生素.Text = lbl抗生素.Tag
            zlControl.TxtSelAll txt抗生素
            Exit Sub
        End If
    End If
    txt抗生素.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = Dkp_ID_vfg Then
        Item.Handle = picVfg.hWnd
    ElseIf Item.ID = Dkp_ID_Rpt Then
        Item.Handle = rptList.hWnd
    ElseIf Item.ID = Dkp_ID_Edit Then
        Item.Handle = picEdit.hWnd
    End If
End Sub

Private Sub Form_Load()
    Call initDockPane
    Call initEdit
    Call initRpt
    
    '内部菜单工具栏定义
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlcommfun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
    End With
    Me.cbsThis.EnableCustomization False
    
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.Position = xtpBarBottom
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With Me.cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加新行"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除本行"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.Style = xtpButtonIconAndCaption
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "更新到参考值列表中"): cbrControl.Flags = xtpFlagRightAlign: cbrControl.Style = xtpButtonIconAndCaption
    End With
    
    If Me.rptList.Tag = "Unload" Then
        MsgBox "请先设置“药敏试验抗生素组”后再使用此功能！", vbInformation, Me.Caption
        Unload Me
    End If
    
End Sub



Private Sub picVfg_Resize()
    With vfgList
        .Left = picVfg.ScaleLeft
        .Width = picVfg.ScaleWidth
        .Height = picVfg.ScaleHeight
        .Top = picVfg.ScaleTop
    End With
End Sub

Private Sub rptList_SelectionChanged()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng分组ID As Long
    
      '展开选中行
    With rptList
        If .FocusedRow Is Nothing And .Rows.count > 0 Then
            If .Rows(0).GroupRow Then
                Set .FocusedRow = .Rows(0).Childs(0)
            Else
                Set .FocusedRow = .Rows(0)
            End If
        End If
        If .FocusedRow Is Nothing Then Exit Sub
    End With
    
    If Not rptList.FocusedRow.GroupRow Then
        lng分组ID = Val(rptList.FocusedRow.Record(mCol.分组Id).Value)
        mlngItemID = lng分组ID
        Call initVfg(lng分组ID)
     End If

End Sub

Private Sub txt参考值_GotFocus(Index As Integer)
    Me.txt参考值(Index).SelStart = 0: Me.txt参考值(Index).SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt参考值_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub initEdit()
    Dim i As Integer
    cbo方法.Clear
    cbo方法.AddItem "MIC", 0
    cbo方法.AddItem "DISK", 1
    cbo方法.AddItem "K-B", 2
    cbo方法.ListIndex = 0
    
    cbo判断方式.Clear
    cbo判断方式.AddItem "参考值除外", 0
    cbo判断方式.AddItem "包含参考值", 1
    cbo判断方式.ListIndex = 1
    
    For i = 0 To 2
        cbo结果(i).Clear
        cbo结果(i).AddItem "R-耐药"
        cbo结果(i).AddItem "I-中介"
        cbo结果(i).AddItem "S-敏感"
        cbo结果(i).ListIndex = 1
    Next
End Sub

Private Sub initDockPane()
    Dim paneRpt As Pane, paneVfg As Pane, paneEdit As Pane
    
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    
    Me.dkpMain.Options.HideClient = True
    
    Set paneRpt = Me.dkpMain.CreatePane(Dkp_ID_Rpt, 90, 190, DockLeftOf, Nothing)
    paneRpt.Title = "抗生素分组"
    paneRpt.Handle = Me.rptList.hWnd
    paneRpt.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set paneVfg = Me.dkpMain.CreatePane(Dkp_ID_vfg, 230, 300, DockRightOf, paneRpt)
    paneVfg.Title = "参考值列表"
    paneVfg.Handle = Me.picVfg.hWnd
    paneVfg.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set paneEdit = Me.dkpMain.CreatePane(Dkp_ID_Edit, 100, 180, DockBottomOf, paneVfg)
    paneEdit.Title = ""
    paneEdit.Handle = Me.picEdit.hWnd
    paneEdit.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
End Sub

Private Sub initRpt()
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo errHandle
    With rptList
        .Columns.DeleteAll
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.分组Id, "分组id", 0, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编号, "编号", 60, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.名称, "分组", 80, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.英文, "英文", 80, True): rptCol.Editable = False: rptCol.Groupable = False
    
        .Records.DeleteAll '清空原列表
        strSQL = "Select B.id,B.编码, B.名称, B.英文 From 检验抗生素组 B, 检验细菌抗生素 A Where A.抗生素分组id = B.ID And A.细菌id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng细菌id)
        If rsTmp.EOF Then
            Me.rptList.Tag = "Unload"
        Else
            Me.rptList.Tag = ""
        End If
        
        Do Until rsTmp.EOF
        
            Set rptRcd = Me.rptList.Records.Add()
            
            Set rptItem = rptRcd.AddItem(""): rptItem.Focusable = False
            Set rptItem = rptRcd.AddItem(CStr("" & rsTmp!ID)): rptItem.Focusable = False
            Set rptItem = rptRcd.AddItem(CStr("" & rsTmp!编码)): rptItem.Focusable = False
            Set rptItem = rptRcd.AddItem(CStr("" & rsTmp!名称)): rptItem.Focusable = False
            Set rptItem = rptRcd.AddItem(CStr("" & rsTmp!英文)): rptItem.Focusable = False
            rsTmp.MoveNext
        Loop
        .Populate
        Call rptList_SelectionChanged '触发选择事件
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub initVfg(ByVal lng分组ID As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngCount As Long
    Dim BlnFind As Boolean
    On Error GoTo errHandle
    With vfgList
        .Rows = 2: .Cols = 14: .FixedRows = 1: .FixedCols = 0
        
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.编码) = "编码"
        .TextMatrix(0, mCol.中文名) = "中文名": .TextMatrix(0, mCol.英文名) = "英文名"
        .TextMatrix(0, mCol.药敏方法) = "药敏方法": .TextMatrix(0, mCol.参考低值) = "参考低值"
        .TextMatrix(0, mCol.参考高值) = "参考高值": .TextMatrix(0, mCol.参考) = "参考"
        .TextMatrix(0, mCol.判断方式) = "判断方式": .TextMatrix(0, mCol.备注) = "备注"
        .TextMatrix(0, mCol.关键字) = "关键字": .TextMatrix(0, mCol.低值结果) = "低值结果"
        .TextMatrix(0, mCol.中间结果) = "中间结果": .TextMatrix(0, mCol.高值结果) = "高值结果"
        
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.编码) = 1000
        .ColWidth(mCol.中文名) = 2600: .ColWidth(mCol.英文名) = 1000: .ColWidth(mCol.药敏方法) = 800
        .ColWidth(mCol.参考低值) = 0: .ColWidth(mCol.参考高值) = 0: .ColWidth(mCol.参考) = 1000
        .ColWidth(mCol.判断方式) = 1000: .ColWidth(mCol.备注) = 1000: .ColWidth(mCol.关键字) = 0
        .ColWidth(mCol.低值结果) = 0: .ColWidth(mCol.中间结果) = 0: .ColWidth(mCol.高值结果) = 0
        
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        
        strSQL = "Select B.ID, B.编码, B.中文名, B.英文名, A.药敏方法, A.参考低值, A.参考高值, A.判断方式, A.备注, A.低值结果, A.中间结果, A.高值结果" & vbNewLine & _
                "From 检验用抗生素 B, 检验细菌抗生素参考 A" & vbNewLine & _
                "Where A.抗生素id = B.ID And A.细菌id = [1] And A.抗生素分组id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng细菌id, lng分组ID)
        
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, mCol.ID) = Val("" & rsTmp!ID)
            .TextMatrix(.Rows - 1, mCol.编码) = "" & rsTmp!编码
            .TextMatrix(.Rows - 1, mCol.中文名) = "" & rsTmp!中文名
            .TextMatrix(.Rows - 1, mCol.英文名) = "" & rsTmp!英文名
            .TextMatrix(.Rows - 1, mCol.药敏方法) = Get药敏方法("" & rsTmp!药敏方法)             '1-MIC;2-DISK;3-K-B
            .TextMatrix(.Rows - 1, mCol.参考低值) = FormatDecimal("" & rsTmp!参考低值)
            .TextMatrix(.Rows - 1, mCol.参考高值) = FormatDecimal("" & rsTmp!参考高值)
            If "" & rsTmp!参考低值 = "" Or "" & rsTmp!参考高值 = "" Then
                .TextMatrix(.Rows - 1, mCol.参考) = FormatDecimal("" & rsTmp!参考低值) & FormatDecimal("" & rsTmp!参考高值)
            Else
                .TextMatrix(.Rows - 1, mCol.参考) = FormatDecimal("" & rsTmp!参考低值) & "～" & FormatDecimal("" & rsTmp!参考高值)
            End If
            
            .TextMatrix(.Rows - 1, mCol.判断方式) = Get判断方式("" & rsTmp!判断方式)
            .TextMatrix(.Rows - 1, mCol.备注) = "" & rsTmp!备注
            .TextMatrix(.Rows - 1, mCol.关键字) = Val("" & rsTmp!ID) & "," & rsTmp!药敏方法
            .TextMatrix(.Rows - 1, mCol.低值结果) = IIf(Trim("" & rsTmp!低值结果) = "", "S-敏感", "" & rsTmp!低值结果)
            .TextMatrix(.Rows - 1, mCol.中间结果) = IIf(Trim("" & rsTmp!中间结果) = "", "I-中介", "" & rsTmp!中间结果)
            .TextMatrix(.Rows - 1, mCol.高值结果) = IIf(Trim("" & rsTmp!高值结果) = "", "R-耐药", "" & rsTmp!高值结果)
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        
        '---- 加上未定义的记录,方便用户操作
        strSQL = "Select A.ID, A.编码, A.中文名, A.英文名, A.药敏方法" & vbNewLine & _
                "From 检验抗生素用药 B, 检验用抗生素 A" & vbNewLine & _
                "Where A.ID = B.抗生素id And B.抗生素分组id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng分组ID)
        Do Until rsTmp.EOF
            BlnFind = False
            For lngCount = .FixedRows To .Rows - 1
                If .TextMatrix(lngCount, mCol.关键字) = Val("" & rsTmp!ID) & "," & rsTmp!药敏方法 Then
                    BlnFind = True
                End If
            Next
            If BlnFind = False Then
                If Val(.TextMatrix(.Rows - 1, mCol.ID)) <> 0 Then
                    .Rows = .Rows + 1
                End If
                .TextMatrix(.Rows - 1, mCol.ID) = Val("" & rsTmp!ID)
                .TextMatrix(.Rows - 1, mCol.编码) = "" & rsTmp!编码
                .TextMatrix(.Rows - 1, mCol.中文名) = "" & rsTmp!中文名
                .TextMatrix(.Rows - 1, mCol.英文名) = "" & rsTmp!英文名
                .TextMatrix(.Rows - 1, mCol.药敏方法) = Get药敏方法("" & rsTmp!药敏方法)
                .TextMatrix(.Rows - 1, mCol.判断方式) = Get判断方式("1")
                .TextMatrix(.Rows - 1, mCol.关键字) = Val("" & rsTmp!ID) & "," & rsTmp!药敏方法
                .TextMatrix(.Rows - 1, mCol.低值结果) = "S-敏感"
                .TextMatrix(.Rows - 1, mCol.中间结果) = "I-中介"
                .TextMatrix(.Rows - 1, mCol.高值结果) = "R-耐药"

            End If
            rsTmp.MoveNext
        Loop
        
       ' If .Rows > 2 Then .Rows = .Rows - 1
        
        .Select .FixedRows, mCol.编号
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function Get判断方式(ByVal strIn) As String
    '0-参考值除外 1-包含参考值
    If strIn = "0" Then
        Get判断方式 = "参考值除外"
    ElseIf strIn = "1" Then
        Get判断方式 = "包含参考值"
    ElseIf strIn = "参考值除外" Then
        Get判断方式 = "0"
    ElseIf strIn = "包含参考值" Then
        Get判断方式 = "1"
    End If
End Function

Private Function Get药敏方法(ByVal strIn) As String
    '1-MIC;2-DISK;3-K-B
    If strIn = "1" Then
        Get药敏方法 = "MIC"
    ElseIf strIn = "2" Then
        Get药敏方法 = "DISK"
    ElseIf strIn = "3" Then
        Get药敏方法 = "K-B"
    ElseIf strIn = "MIC" Then
        Get药敏方法 = "1"
    ElseIf strIn = "DISK" Then
        Get药敏方法 = "2"
    ElseIf strIn = "K-B" Then
        Get药敏方法 = "3"
    End If
    
End Function

Public Sub ShowMe(ByVal lng细菌ID As Long, ByVal frmMain As Form)

    If lng细菌ID <= 0 Then Exit Sub
    mlng细菌id = lng细菌ID
    On Error Resume Next
    Me.Show vbModal, frmMain
    
End Sub

Private Sub txt参考值_Validate(Index As Integer, Cancel As Boolean)
    txt参考值(Index).Text = FormatDecimal(txt参考值(Index).Text)
End Sub

Private Function FormatDecimal(ByVal strIn As String) As String
    '将.5这种文本格式为0.5
    Dim strTmp As String
    If InStr(strIn, ".") > 0 Then
        strTmp = Mid(strIn, InStr(strIn, ".") + 1)
        FormatDecimal = Format(strIn, "0." & String(Len(strTmp), "0"))
    Else
        FormatDecimal = strIn
    End If
End Function

Private Sub txt抗生素_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt抗生素.Text = "" And txt抗生素.Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    strText = txt抗生素.Text
    
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
        
    gstrSql = "Select B.ID, B.编码, B.中文名, B.英文名, 药敏方法" & vbNewLine & _
            "From 检验用抗生素 B, 检验抗生素用药 A" & vbNewLine & _
            "Where A.抗生素id = B.Id  And A.抗生素分组id = [1] And (" & _
            zlcommfun.GetLike("B", "编码", strText) & " or " & zlcommfun.GetLike("B", "中文名", strText) & " or " & zlcommfun.GetLike("B", "简码", strText) & ")"

    'Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID)
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "抗生素选择", True, "", "请选择抗生素", False, False, False, 0, 0, 0, blnReturn, True, False, mlngItemID)
    If blnReturn = False Then
        If rsTemp.EOF = True Then
            '记录集中没有可选择的数据
            txt抗生素.Text = lbl抗生素.Tag
            zlControl.TxtSelAll txt抗生素
            Exit Sub
        Else
            '肯定是有记录集的
            txt抗生素.Tag = rsTemp!ID
            txt抗生素.Text = "(" & rsTemp!编码 & ")" & rsTemp!中文名
            lbl抗生素.Tag = txt抗生素.Text '用于恢复显示
        End If
    End If
    cbo方法.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_RowColChange()
    
    Dim str方法  As String
    Dim str判断 As String
    Dim str结果 As String
    With vfgList
        If .Row >= .FixedRows Then
            If Val(.TextMatrix(.Row, mCol.ID)) = 0 Then Exit Sub
            txt抗生素.Tag = Val(.TextMatrix(.Row, mCol.ID))
            txt抗生素.Text = "(" & .TextMatrix(.Row, mCol.编码) & ")" & .TextMatrix(.Row, mCol.中文名)
            lbl抗生素.Tag = txt抗生素.Text
            str方法 = .TextMatrix(.Row, mCol.药敏方法)
            str判断 = .TextMatrix(.Row, mCol.判断方式)
            
            cbo方法.ListIndex = Val(Get药敏方法(str方法)) - 1
            cbo判断方式.ListIndex = Val(Get判断方式(str判断))
            txt参考值(0) = .TextMatrix(.Row, mCol.参考低值)
            txt参考值(1) = .TextMatrix(.Row, mCol.参考高值)
            txt备注 = .TextMatrix(.Row, mCol.备注)
            
            str结果 = .TextMatrix(.Row, mCol.低值结果)
            cbo结果(0).Text = str结果
            str结果 = .TextMatrix(.Row, mCol.中间结果)
            cbo结果(1).Text = str结果
            str结果 = .TextMatrix(.Row, mCol.高值结果)
            cbo结果(2).Text = str结果
            
        End If
    End With
End Sub
