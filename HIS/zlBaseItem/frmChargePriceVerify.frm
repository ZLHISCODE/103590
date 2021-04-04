VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChargePriceVerify 
   Caption         =   "调价审核"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11100
   Icon            =   "frmChargePriceVerify.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11100
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退出(&C)"
      Height          =   350
      Left            =   8760
      TabIndex        =   3
      Tag             =   "分类"
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "审核(&O)"
      Height          =   350
      Left            =   7470
      TabIndex        =   2
      Tag             =   "分类"
      Top             =   6480
      Width           =   1100
   End
   Begin TabDlg.SSTab ssTdetails 
      Height          =   6255
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "未审核"
      TabPicture(0)   =   "frmChargePriceVerify.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vsfNotList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "已审核"
      TabPicture(1)   =   "frmChargePriceVerify.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDateArea"
      Tab(1).Control(1)=   "lblTo"
      Tab(1).Control(2)=   "vsfList"
      Tab(1).Control(3)=   "cobDateArea"
      Tab(1).Control(4)=   "dtpDateBegin"
      Tab(1).Control(5)=   "dtpDateEnd"
      Tab(1).Control(6)=   "cmdFilter"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdFilter 
         Caption         =   "过滤(&F)"
         Height          =   300
         Left            =   -69000
         TabIndex        =   9
         Top             =   420
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpDateEnd 
         Height          =   300
         Left            =   -70680
         TabIndex        =   8
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   158859265
         CurrentDate     =   42067
      End
      Begin MSComCtl2.DTPicker dtpDateBegin 
         Height          =   300
         Left            =   -72480
         TabIndex        =   6
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   158859265
         CurrentDate     =   42067
      End
      Begin VB.ComboBox cobDateArea 
         Height          =   300
         Left            =   -74080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   420
         Width           =   1455
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfNotList 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8175
         _cx             =   14420
         _cy             =   7858
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmChargePriceVerify.frx":688A
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
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   10
         Top             =   840
         Width           =   8175
         _cx             =   14420
         _cy             =   7223
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmChargePriceVerify.frx":6A81
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
         VirtualData     =   0   'False
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
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         Caption         =   "~"
         Height          =   180
         Left            =   -70920
         TabIndex        =   7
         Top             =   480
         Width           =   90
      End
      Begin VB.Label lblDateArea 
         AutoSize        =   -1  'True
         Caption         =   "日期范围"
         Height          =   180
         Left            =   -74950
         TabIndex        =   4
         Top             =   480
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmChargePriceVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnCanUpdateAll As Boolean '是否允许操作所有项目：未启用价格等级或启用了价格等级有“所有院区”权限

Public Sub ShowMe(ByVal frmParent As Form, ByVal blnCanUpdateAll As Boolean)
    '公共函数，打开窗体
    mblnCanUpdateAll = blnCanUpdateAll
    
    Me.Show vbModal, frmParent
End Sub

Private Sub InitComBox()
    '初始化下拉列表
    With cobDateArea
        .AddItem "一个月内"
        .AddItem "三个月内"
        .AddItem "半年内"
        .AddItem "自定义"
        
        .ListIndex = 0
    End With
End Sub

Private Sub GetNotVerifyData()
    '获取未审核数据
    Dim rsData As ADODB.Recordset
    Dim strWhere As String
    
    On Error GoTo ErrHandle
    If mblnCanUpdateAll = False Then
        strWhere = " And (c.站点=[1]" & vbNewLine & _
                "       Or c.站点 Is Null And a.价格等级 In(" & vbNewLine & _
                "           Select m.名称" & vbNewLine & _
                "           From 收费价格等级 M, 收费价格等级应用 N" & vbNewLine & _
                "           Where m.名称 = n.价格等级 And Nvl(m.是否适用普通项目, 0) = 1 And n.站点 = [1]" & vbNewLine & _
                "                 And (m.撤档时间 Is Null Or m.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))))"
    End If
    vsfNotList.Rows = 1
    gstrSQL = "Select a.Id, a.调价id, a.审核标志, c.名称 As 收费细目, b.名称 As 收入项目," & vbNewLine & _
            "       a.原价, a.现价, a.缺省价格, a.填制人, a.填制日期, a.执行日期, a.序号," & vbNewLine & _
            "       a.说明,Nvl(a.价格等级,'缺省') As 价格等级" & vbNewLine & _
            "From 收费调价记录 A, 收入项目 B, 收费项目目录 C" & vbNewLine & _
            "Where a.收入项目id = b.Id And a.收费细目id = c.Id And 审核标志 = 0" & strWhere & vbNewLine & _
            "Order By a.Id, a.调价id, a.序号"
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "未审核单据查询", gstrNodeNo)
    With vsfNotList
        .MergeCells = flexMergeRestrictColumns
        Do While Not rsData.EOF
            .MergeCol(.ColIndex("收费细目")) = True   '与上面的.MergeCells属性结合使用不同行同列内容相同的合并
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsData!ID
            .TextMatrix(.Rows - 1, .ColIndex("调价id")) = rsData!调价id
            .TextMatrix(.Rows - 1, .ColIndex("审核状态")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("收费细目")) = rsData!收费细目
            .TextMatrix(.Rows - 1, .ColIndex("价格等级")) = Nvl(rsData!价格等级)
            .TextMatrix(.Rows - 1, .ColIndex("收入项目")) = rsData!收入项目
            .TextMatrix(.Rows - 1, .ColIndex("原价")) = IIF(IsNull(rsData!原价), "", rsData!原价)
            .TextMatrix(.Rows - 1, .ColIndex("现价")) = IIF(IsNull(rsData!现价), "", rsData!现价)
            .TextMatrix(.Rows - 1, .ColIndex("缺省价格")) = IIF(IsNull(rsData!缺省价格), "", rsData!缺省价格)
            .TextMatrix(.Rows - 1, .ColIndex("填制人")) = IIF(IsNull(rsData!填制人), "", rsData!填制人)
            .TextMatrix(.Rows - 1, .ColIndex("填制日期")) = IIF(IsNull(rsData!填制日期), "", Format(rsData!填制日期, "yyyy-mm-dd hh:mm:ss"))
            .TextMatrix(.Rows - 1, .ColIndex("执行日期")) = IIF(IsNull(rsData!执行日期), "", Format(rsData!执行日期, "yyyy-mm-dd hh:mm:ss"))
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = rsData!序号
            .TextMatrix(.Rows - 1, .ColIndex("说明")) = IIF(IsNull(rsData!说明), "", rsData!说明)
            
            rsData.MoveNext
        Loop
        If .Rows > 1 Then
            .Cell(flexcpBackColor, 1, .ColIndex("审核状态"), .Rows - 1, .ColIndex("审核状态")) = vbWhite
            .Cell(flexcpBackColor, 1, .ColIndex("说明"), .Rows - 1, .ColIndex("说明")) = vbWhite
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetVerifyData()
    '获取已经审核数据
    Dim rsData As ADODB.Recordset
    Dim datBeginDate As Date
    Dim datEndDate As Date
    Dim strWhere As String
    
    On Error GoTo ErrHandle
    If mblnCanUpdateAll = False Then
        strWhere = " And (c.站点=[3]" & vbNewLine & _
                "       Or c.站点 Is Null And a.价格等级 In(" & vbNewLine & _
                "           Select m.名称" & vbNewLine & _
                "           From 收费价格等级 M, 收费价格等级应用 N" & vbNewLine & _
                "           Where m.名称 = n.价格等级 And Nvl(m.是否适用普通项目, 0) = 1 And n.站点 = [3]" & vbNewLine & _
                "                 And (m.撤档时间 Is Null Or m.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))))"
    End If
    '已审核单据查询
    If cobDateArea.Text <> "自定义" Then
        Select Case cobDateArea.Text
        Case "一个月内"
            datBeginDate = CDate(Format(DateAdd("M", -1, Date), "yyyy-mm-dd") & " 00:00:00")
            datEndDate = sys.Currentdate
        Case "三个月内"
            datBeginDate = CDate(Format(DateAdd("M", -3, Date), "yyyy-mm-dd") & " 00:00:00")
            datEndDate = sys.Currentdate
        Case "半年内"
            datBeginDate = CDate(Format(DateAdd("M", -6, Date), "yyyy-mm-dd") & " 00:00:00")
            datEndDate = sys.Currentdate
        End Select
    Else
        datBeginDate = CDate(Format(dtpDateBegin, "yyyy-mm-dd") & " 00:00:00")
        datEndDate = CDate(Format(dtpDateEnd, "yyyy-mm-dd") & " 23:59:59")
    End If
    gstrSQL = "Select a.Id, a.调价id, a.审核标志, c.名称 As 收费细目, b.名称 As 收入项目," & vbNewLine & _
            "       a.原价, a.现价, a.缺省价格, a.填制人, a.填制日期, a.执行日期, a.序号," & vbNewLine & _
            "       a.审核人, a.审核日期,a.说明,Nvl(a.价格等级,'缺省') As 价格等级" & vbNewLine & _
            "From 收费调价记录 A, 收入项目 B, 收费项目目录 C" & vbNewLine & _
            "Where a.收入项目id = b.Id And a.收费细目id = c.Id And (审核标志 = 1 Or 审核标志 = 2)" & vbNewLine & _
            "       And a.审核日期 Between [1] And [2]" & strWhere & vbNewLine & _
            "Order By a.Id, a.调价id, a.序号"
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "已审核单据查询", datBeginDate, datEndDate, gstrNodeNo)
    With VSFList
        .Rows = 1
        Do While Not rsData.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsData!ID
            .TextMatrix(.Rows - 1, .ColIndex("调价id")) = rsData!调价id
            If rsData!审核标志 = 1 Then
                .TextMatrix(.Rows - 1, .ColIndex("审核状态")) = "√"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("审核状态")) = "×"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("收费细目")) = rsData!收费细目
            .TextMatrix(.Rows - 1, .ColIndex("价格等级")) = Nvl(rsData!价格等级)
            .TextMatrix(.Rows - 1, .ColIndex("收入项目")) = rsData!收入项目
            .TextMatrix(.Rows - 1, .ColIndex("原价")) = IIF(IsNull(rsData!原价), "", rsData!原价)
            .TextMatrix(.Rows - 1, .ColIndex("现价")) = IIF(IsNull(rsData!现价), "", rsData!现价)
            .TextMatrix(.Rows - 1, .ColIndex("缺省价格")) = IIF(IsNull(rsData!缺省价格), "", rsData!缺省价格)
            .TextMatrix(.Rows - 1, .ColIndex("填制人")) = IIF(IsNull(rsData!填制人), "", rsData!填制人)
            .TextMatrix(.Rows - 1, .ColIndex("填制日期")) = IIF(IsNull(rsData!填制日期), "", Format(rsData!填制日期, "yyyy-mm-dd hh:mm:ss"))
            .TextMatrix(.Rows - 1, .ColIndex("执行日期")) = IIF(IsNull(rsData!执行日期), "", Format(rsData!执行日期, "yyyy-mm-dd hh:mm:ss"))
            .TextMatrix(.Rows - 1, .ColIndex("审核人")) = IIF(IsNull(rsData!审核人), "", rsData!审核人)
            .TextMatrix(.Rows - 1, .ColIndex("审核日期")) = IIF(IsNull(rsData!审核日期), "", Format(rsData!审核日期, "yyyy-mm-dd hh:mm:ss"))
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = rsData!序号
            .TextMatrix(.Rows - 1, .ColIndex("说明")) = IIF(IsNull(rsData!说明), "", rsData!说明)
            
            rsData.MoveNext
        Loop
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFilter_Click()
    Call GetVerifyData
End Sub

Private Sub cmdOK_Click()
    Dim int审核标志 As Integer
    Dim intRow As Integer
    Dim str说明 As String
    Dim rsUser As New ADODB.Recordset
    Dim str用户名 As String
    Dim str审核日期 As String
        
    With vsfNotList
        If .Rows > 1 Then
            Set rsUser = sys.GetUserInfo
            str用户名 = IIF(IsNull(rsUser!姓名), "", rsUser!姓名) '当前用户姓名
            str审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, .ColIndex("审核状态")) <> "" Then
                    Select Case .TextMatrix(intRow, .ColIndex("审核状态"))
                    Case "√"
                        int审核标志 = 1
                    Case "×"
                        int审核标志 = 2
                    End Select
                    str说明 = .TextMatrix(intRow, .ColIndex("说明"))
                    gstrSQL = "Zl_收费调价记录_Verify(" & _
                    .TextMatrix(intRow, .ColIndex("id")) & "," & _
                    int审核标志 & ",'" & _
                    str用户名 & "'," & _
                    "to_date('" & Format(str审核日期, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS') " & _
                    IIF(str说明 = "", ")", ",'" & str说明 & "')")
                
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "审核单据")
                End If
            Next
            
            Call GetNotVerifyData
        End If
    End With
End Sub

Private Sub cobDateArea_Click()
    With cobDateArea
        If .Text = "自定义" Then
            dtpDateBegin.Visible = True
            dtpDateEnd.Visible = True
            lblTo.Visible = True
            
            dtpDateBegin.Move cobDateArea.Left + cobDateArea.Width + 100
            lblTo.Move dtpDateBegin.Left + dtpDateBegin.Width + 100
            dtpDateEnd.Move lblTo.Left + lblTo.Width + 100
            cmdFilter.Move dtpDateEnd.Left + dtpDateEnd.Width + 100
        Else
            dtpDateBegin.Visible = False
            dtpDateEnd.Visible = False
            lblTo.Visible = False
            
            cmdFilter.Move cobDateArea.Left + cobDateArea.Width + 100
        End If
    End With
End Sub

Private Sub Form_Load()
    Call InitComBox
    Call GetNotVerifyData
    Call GetVerifyData
    
    dtpDateEnd = sys.Currentdate
    dtpDateBegin = DateAdd("m", -1, dtpDateEnd)
End Sub

Private Sub Form_Resize()
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 100, Me.ScaleHeight - cmdCancel.Height - 100
    cmdOK.Move cmdCancel.Left - cmdOK.Width - 100, cmdCancel.Top
    ssTdetails.Move 40, 0, Me.ScaleWidth - 100, Me.ScaleHeight - cmdOK.Height - 300
    vsfNotList.Move 10, 400, ssTdetails.Width - 20, ssTdetails.Height - 380
    VSFList.Move 35, cobDateArea.Top + cobDateArea.Height + 50, ssTdetails.Width - 20, ssTdetails.Height - VSFList.Top + 80
End Sub

Private Sub ssTdetails_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        cmdOK.Visible = False
        vsfNotList.Visible = False
        VSFList.Visible = True
        
        Call GetVerifyData
    Else
        cmdOK.Visible = True
        vsfNotList.Visible = True
        VSFList.Visible = False
    End If
End Sub

Private Sub vsfList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '移动第一栏的标记到当前行！
    With VSFList
        .Cell(flexcpText, 0, 0, .Rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 8
        End If
    End With
End Sub

Private Sub vsfNotList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '移动第一栏的标记到当前行！
    With vsfNotList
        .Cell(flexcpText, 0, 0, .Rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 8
        End If
    End With
End Sub

Private Sub vsfNotList_DblClick()
    Dim intRow As Integer
    
    With vsfNotList
        If .Row > 0 And .Col = .ColIndex("审核状态") Then
            If .TextMatrix(.Row, .ColIndex("审核状态")) = "" Then
                .TextMatrix(.Row, .ColIndex("审核状态")) = "√"
                For intRow = 1 To .Rows - 1
                    If .TextMatrix(.Row, .ColIndex("调价id")) = .TextMatrix(intRow, .ColIndex("调价id")) Then
                        .TextMatrix(intRow, .ColIndex("审核状态")) = "√"
                    End If
                Next
            ElseIf .TextMatrix(.Row, .ColIndex("审核状态")) = "√" Then
                .TextMatrix(.Row, .ColIndex("审核状态")) = "×"
                For intRow = 1 To .Rows - 1
                    If .TextMatrix(.Row, .ColIndex("调价id")) = .TextMatrix(intRow, .ColIndex("调价id")) Then
                        .TextMatrix(intRow, .ColIndex("审核状态")) = "×"
                    End If
                Next
            Else
                .TextMatrix(.Row, .ColIndex("审核状态")) = ""
                For intRow = 1 To .Rows - 1
                    If .TextMatrix(.Row, .ColIndex("调价id")) = .TextMatrix(intRow, .ColIndex("调价id")) Then
                        .TextMatrix(intRow, .ColIndex("审核状态")) = ""
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub vsfNotList_EnterCell()
    With vsfNotList
        If .Rows > 1 Then
            If .Col = .ColIndex("说明") Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End If
    End With
End Sub
