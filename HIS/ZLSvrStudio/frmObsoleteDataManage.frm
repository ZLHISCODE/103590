VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmObsoleteDataManage 
   BackColor       =   &H80000005&
   Caption         =   "历史数据空间管理"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmObsoleteDataManage.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   8325
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid vsfBusinessList 
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   1695
      Width           =   7755
      _cx             =   13679
      _cy             =   4048
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   900
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmObsoleteDataManage.frx":04F9
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
      ExplorerBar     =   1
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
   Begin VB.CommandButton cmdFunction 
      Caption         =   "数据处理(&D)"
      Height          =   350
      Index           =   1
      Left            =   6660
      TabIndex        =   3
      Top             =   1245
      Width           =   1200
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "数据查询(&S)"
      Height          =   350
      Index           =   0
      Left            =   5445
      TabIndex        =   2
      Top             =   1245
      Width           =   1200
   End
   Begin MSComctlLib.ImageList imgSys 
      Left            =   5280
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObsoleteDataManage.frx":0617
            Key             =   "Other"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObsoleteDataManage.frx":16A9
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObsoleteDataManage.frx":4A9B
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObsoleteDataManage.frx":7E8D
            Key             =   "LockAndRun"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMain 
      BackColor       =   &H8000000E&
      Height          =   15
      Left            =   120
      TabIndex        =   5
      Top             =   5250
      Width           =   6360
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "未完结业务清单"
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   1410
      Width           =   1935
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmObsoleteDataManage.frx":B27F
      Height          =   540
      Left            =   855
      TabIndex        =   1
      Top             =   660
      Width           =   8460
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "未完业务数据管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   1920
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   165
      Picture         =   "frmObsoleteDataManage.frx":B38B
      Top             =   675
      Width           =   480
   End
End
Attribute VB_Name = "frmObsoleteDataManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const vsfEditBackColor = &HFFF0D2       '方案内容表格标题背景颜色
Private mblnOwner As Boolean

Private Enum cmdFunc
    FCT_数据查询 = 0
    FCT_数据处理 = 1
End Enum

Private Enum BusinessList
    BL_系统编号 = 0
    BL_系统 = 1
    BL_名称 = 2
    BL_是否定期处理 = 3
    BL_保留天数 = 4
    BL_数据处理过程 = 5
    BL_最后操作人员 = 6
    BL_最后操作时间 = 7
    BL_说明 = 8
End Enum

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Private Sub cmdFunction_Click(Index As Integer)
    Dim lngDays As Long
    Dim strUserName As String
    Dim dateNow As Date
    
    If CheckUser = False Then Exit Sub
    With vsfBusinessList
        Select Case Index
            Case FCT_数据查询
                Call frmObsoleteDataQuery.ShowMe(.TextMatrix(.RowSel, BL_名称))
            Case FCT_数据处理
                lngDays = .TextMatrix(.RowSel, BL_保留天数)
                strUserName = .TextMatrix(.RowSel, BL_最后操作人员)
                If frmObsoleteDataDeal.ShowMe(.TextMatrix(.RowSel, BL_名称), .TextMatrix(.RowSel, BL_数据处理过程), lngDays, strUserName, dateNow) Then
                    If lngDays <> .TextMatrix(.RowSel, BL_保留天数) Then
                         .TextMatrix(.RowSel, BL_保留天数) = lngDays
                    End If
                    .TextMatrix(.RowSel, BL_最后操作时间) = Format(dateNow, "yyyy-MM-dd HH:mm:ss")
                End If
        End Select
    End With
End Sub

Private Sub Form_Activate()
    If vsfBusinessList.Rows = 1 Then
        vsfBusinessList.Rows = 2
        cmdFunction(FCT_数据查询).Enabled = False
        cmdFunction(FCT_数据处理).Enabled = False
        MsgBox "当前用户为：" & gstrLoginUserName & "，该用户下没有“未完结业务”。"
    End If
End Sub

Private Sub Form_Load()
    '填充未完结业务清单
    Call FillData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsfBusinessList.Width = Me.ScaleWidth - vsfBusinessList.Left * 2
    vsfBusinessList.Height = Me.ScaleHeight - vsfBusinessList.Top - 50
    cmdFunction(FCT_数据查询).Left = Me.ScaleWidth - cmdFunction(FCT_数据查询).Width * 2 - 30 - 145
    Call SetCtrlPosOnLine(False, 0, cmdFunction(FCT_数据查询), 15, cmdFunction(FCT_数据处理))
    If vsfBusinessList.Rows > 1 Then
        vsfBusinessList.Cell(flexcpBackColor, 1, BL_是否定期处理, vsfBusinessList.Rows - 1, BL_保留天数) = vsfEditBackColor
    End If
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub FillData()
'填充未完结业务清单
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    '判断gstrUserName是否为系统所有者
    strSQL = "SELECT 1 FROM ZLTOOLS.ZLSYSTEMS WHERE 所有者=USER"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "系统所有者判定")
    mblnOwner = Not rsTemp.EOF
    strSQL = "Select b.编号, b.名称 系统名称, a.名称, a.说明, a.是否定期处理, a.保留天数, a.数据处理过程, a.操作人员, To_char(a.操作时间,'yyyy-MM-dd hh24:mi:ss') 操作时间" & vbNewLine & _
            "From Zltools.Zlobsoletedatadeal A, zlSystems B" & vbNewLine & _
            "Where a.系统 = b.编号 " & IIf(mblnOwner, "and b.所有者 = [1]", "Order by a.系统")
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取未完结业务清单", gstrUserName)
    With rsTemp
        vsfBusinessList.Redraw = flexRDNone
        vsfBusinessList.Rows = vsfBusinessList.FixedRows
        vsfBusinessList.Rows = .RecordCount + 1
        vsfBusinessList.WordWrap = True
        
        For i = 1 To .RecordCount
            vsfBusinessList.TextMatrix(i, BL_系统编号) = !编号
            vsfBusinessList.TextMatrix(i, BL_系统) = !系统名称
            vsfBusinessList.TextMatrix(i, BL_名称) = !名称
            vsfBusinessList.TextMatrix(i, BL_是否定期处理) = IIf(!是否定期处理 = 1, "√", "×")
            vsfBusinessList.TextMatrix(i, BL_保留天数) = Nvl(!保留天数, 0)
            vsfBusinessList.TextMatrix(i, BL_数据处理过程) = !数据处理过程
            vsfBusinessList.TextMatrix(i, BL_最后操作人员) = !操作人员 & ""
            vsfBusinessList.TextMatrix(i, BL_最后操作时间) = !操作时间 & ""
            vsfBusinessList.TextMatrix(i, BL_说明) = !说明 & ""
            .MoveNext
        Next
        vsfBusinessList.Redraw = flexRDDirect
        vsfBusinessList.AutoSize BL_系统编号, BL_说明
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub vsfBusinessList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dateNow As Date

    On Error GoTo errH
    With vsfBusinessList
        .Text = Val(.Text)
        .Tag = Val(.Tag)
        
        If Col = BL_保留天数 And .Tag <> .Text Then
            If Val(.Text) < 1 Then
                .Text = .Tag
                MsgBox "保留天数至少为1天，请重新调整！", vbInformation, gstrSysName
                Exit Sub
            End If
            dateNow = CurrentDate()
            '更新保留天数信息
            Call ExecuteProcedure("Zltools.Zl_Zlobsoletedatadeal_Update('" & .TextMatrix(.Row, BL_名称) & "',Null," & _
                                                                        .Text & ",'" & _
                                                                        gstrLoginUserName & "','" & _
                                                                        dateNow & "')", "修改保留天数及人员时间信息")
            .TextMatrix(.Row, BL_最后操作人员) = gstrLoginUserName
            .TextMatrix(.Row, BL_最后操作时间) = Format(dateNow, "yyyy-MM-dd HH:mm:ss")
            '插入重要操作日志
            Call SaveAuditLog(2, "修改保留天数", "将业务“" & .TextMatrix(.Row, BL_名称) & "”的数据的保留天数由" & .Tag & "天修改为" & .Text & "天")
        End If
    End With
    Exit Sub
errH:
    vsfBusinessList.Text = vsfBusinessList.Tag
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub vsfBusinessList_DblClick()
    Dim dateNow As Date

    On Error GoTo errH
    With vsfBusinessList
        If .MouseRow <> .Row Then Exit Sub
        If .Row = 0 Then Exit Sub
        '双击“是否定期处理”列，修改状态
        If .ColSel = BL_是否定期处理 Then
            If CheckUser = False Then Exit Sub
            dateNow = CurrentDate()
            
            If .TextMatrix(.RowSel, BL_是否定期处理) = "√" Then
                Call ExecuteProcedure("Zltools.Zl_Zlobsoletedatadeal_Update('" & .TextMatrix(.Row, BL_名称) & _
                                                                            "',0," & _
                                                                            .TextMatrix(.Row, BL_保留天数) & ",'" & _
                                                                            gstrLoginUserName & "','" & _
                                                                            dateNow & "')", "关闭定期处理及修改人员时间信息")
                .TextMatrix(.RowSel, BL_是否定期处理) = "×"
                
                '插入重要操作日志
                Call SaveAuditLog(2, "是否定期处理", "停用业务“" & .TextMatrix(.Row, BL_名称) & "”的数据定期处理")
            ElseIf .TextMatrix(.RowSel, BL_是否定期处理) = "×" Then
                Call ExecuteProcedure("Zltools.Zl_Zlobsoletedatadeal_Update('" & .TextMatrix(.Row, BL_名称) & _
                            "',1," & _
                            .TextMatrix(.Row, BL_保留天数) & ",'" & _
                            gstrLoginUserName & "','" & _
                            dateNow & "')", "开启定期处理及修改人员时间信息")
                .TextMatrix(.RowSel, BL_是否定期处理) = "√"
                
                '插入重要操作日志
                Call SaveAuditLog(2, "是否定期处理", "启用业务“" & .TextMatrix(.Row, BL_名称) & "”的数据定期处理")
            End If
            .TextMatrix(.Row, BL_最后操作人员) = gstrLoginUserName
            .TextMatrix(.Row, BL_最后操作时间) = Format(dateNow, "yyyy-MM-dd HH:mm:ss")
        ElseIf .ColSel = BL_保留天数 Then
            If CheckUser = False Then Exit Sub
            .Tag = .TextMatrix(.Row, BL_保留天数)
            .EditCell
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function CheckUser() As Boolean
    '根据用户身份弹出提示
    If Not mblnOwner Or gstrLoginUserName <> gstrUserName Then
        MsgBox "当前登录用户不是系统所有者用户，故无法执行此操作，" & vbNewLine & "请使用系统所有者用户（" & gstrUserName & "）登录后再执行此操作！"
        Exit Function
    End If
    CheckUser = True
End Function

Private Sub vsfBusinessList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = BL_保留天数 Then
        If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        ElseIf KeyAscii = 13 Then
            Row = Row
        End If
    End If
End Sub

Private Sub vsfBusinessList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    With vsfBusinessList
        If .MouseRow > .Rows - 1 Or .MouseRow <= 0 Then
            Call ShowTipInfo(.hwnd, "")
            
        ElseIf .MouseCol = BL_说明 Then
            Call ShowTipInfo(.hwnd, .TextMatrix(.MouseRow, BL_说明), True)
        Else
            Call ShowTipInfo(.hwnd, "")
        End If
    End With
End Sub
