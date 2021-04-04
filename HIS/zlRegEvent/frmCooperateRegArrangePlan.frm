VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCooperateRegArrangePlan 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicUnit 
      BorderStyle     =   0  'None
      Height          =   6765
      Left            =   0
      ScaleHeight     =   6765
      ScaleWidth      =   2580
      TabIndex        =   3
      Top             =   240
      Width           =   2580
      Begin VB.ListBox lstUnits 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2910
         ItemData        =   "frmCooperateRegArrangePlan.frx":0000
         Left            =   0
         List            =   "frmCooperateRegArrangePlan.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblUnitTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "合作单位"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1020
      End
   End
   Begin VB.PictureBox picUnitReg 
      BorderStyle     =   0  'None
      Height          =   7485
      Left            =   3720
      ScaleHeight     =   7485
      ScaleWidth      =   5580
      TabIndex        =   0
      Top             =   0
      Width           =   5580
      Begin VB.CheckBox chkDisable 
         Caption         =   "本合作单位禁用该号别"
         Height          =   330
         Left            =   2460
         TabIndex        =   6
         Top             =   -15
         Width           =   2265
      End
      Begin VSFlex8Ctl.VSFlexGrid vsUnits 
         Height          =   4455
         Left            =   -240
         TabIndex        =   1
         Top             =   1920
         Width           =   7335
         _cx             =   12938
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCooperateRegArrangePlan.frx":0004
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   110
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
         Editable        =   2
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
      Begin VB.Label lblUnitRegTitle 
         Caption         =   "***:序号分配"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmCooperateRegArrangePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mlngPriItem As Long
Private mlng计划Id              As Long
Private mrs限号                 As ADODB.Recordset
Private mrs计划                 As ADODB.Recordset
Private mstr排班                As String '周日|全日||周一|白天||…………
Private mblnUnload As Boolean
Private mbln时段                As Boolean '如果安排设置了时段则严格按照时段来分配
Private mrs时间段               As ADODB.Recordset
Private mstrKey      As String
Private mrsSource    As ADODB.Recordset
Private mrsUnitsReg  As ADODB.Recordset
Private mrsUnitsInfo As ADODB.Recordset
Private mrsUnits As ADODB.Recordset
Private mrsDisable As ADODB.Recordset
Private mblnChange   As Boolean
Private mbln序号控制 As Boolean
Public Event frmUnload(ByVal blnCancel As Boolean)
Private Sub cmdCancel_Click()
    RaiseEvent frmUnload(True)
End Sub

Private Sub Form_Resize()
   Err.Number = 0
     On Error Resume Next
     If mbln时段 Then
        With Me.PicUnit
            .Left = Me.ScaleLeft
            .Top = Me.ScaleTop
            .Height = Me.ScaleHeight
        End With
        
        With Me.picUnitReg
            .Left = PicUnit.Left + PicUnit.Width + 1 * Screen.TwipsPerPixelX
            .Top = Me.ScaleTop
            .Height = Me.ScaleHeight
            .Width = Me.ScaleWidth - .Left
        End With
     Else
        PicUnit.Visible = False
        With Me.picUnitReg
            .Left = Me.ScaleLeft
            .Top = Me.ScaleTop
            .Height = Me.ScaleHeight
            .Width = Me.ScaleWidth
        End With
     End If
End Sub
 
Public Function frmInit(ByVal lng计划ID As Long, _
    ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '返回:设置成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-29 14:16:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mlng计划Id = lng计划ID
    If InitData() = False Then Exit Function
    mbln时段 = chkExists时段(lng计划ID)
    Call InitRs
    Call InitPage
    If InitUntils() = False Then Exit Function
    If Not mbln时段 Then LoadUnitsReg
   ' Call InitPlan
    frmInit = True
End Function

Private Sub lstUnits_Click()
    Static strUnits As String
    If lstUnits.Text = strUnits Then Exit Sub
    If mblnChange Then
        MoveUnitReg strUnits
    End If
    strUnits = lstUnits.Text
    lblUnitRegTitle.Caption = strUnits & ":合作单位预约分配"
    LoadUnitsReg
    
End Sub

Private Sub LoadUnitsReg()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载已经已经分配的数据信息
    '日期:2013-10-29 18:14:15
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varCol As Variant, i As Long, j        As Long
    Dim lng分配数量 As Long, lngRow As Long
    Dim strUnit As String, str星期 As String
    
    If mbln时段 Then
        If Not mrsDisable Is Nothing Then
            With mrsDisable
                .Filter = "合作单位='" & lstUnits.Text & "'"
                If .RecordCount = 0 Then
                    chkDisable.Value = 0
                Else
                    chkDisable.Value = 1
                End If
            End With
        End If
          With vsUnits
            .Clear 1
            .Rows = 3
            varCol = Split(mstr排班, "||")
            For i = 0 To UBound(varCol)
                mrsSource.Filter = "限制项目='" & Split(varCol(i), "|")(0) & "'"
                 If mrsSource.RecordCount > 0 Then
                    lngRow = 2
                    Do While Not mrsSource.EOF
                        If lngRow + 1 >= .Rows Then .Rows = .Rows + 1
                        .TextMatrix(lngRow, i * 3 + 0) = Nvl(mrsSource!时间段)
                        .TextMatrix(lngRow, i * 3 + 1) = Val(Nvl(mrsSource!数量))
                        mrsUnitsReg.Filter = "限制项目='" & Split(varCol(i), "|")(0) & "' and 合作单位='" & lstUnits.Text & "' and 序号=" & Val(Nvl(mrsSource!序号))
                        If mrsUnitsReg.RecordCount > 0 Then
                            lng分配数量 = Val(Nvl(mrsUnitsReg!数量))
                        Else
                            lng分配数量 = 0
                        End If
                        .TextMatrix(lngRow, i * 3 + 2) = lng分配数量
                        lngRow = lngRow + 1
                        mrsSource.MoveNext
                    Loop
                    
                End If
            Next
            mrsUnitsReg.Filter = 0
            mrsSource.Filter = 0
         End With
         Exit Sub
    End If
    varCol = Split(mstr排班, "||")
    With vsUnits
        For i = 2 To .Rows - 1
            strUnit = .TextMatrix(i, 0)
             If strUnit <> "" Then
                For j = 0 To UBound(varCol)
                   str星期 = Split(varCol(j), "|")(0)
                   mrsUnitsReg.Filter = "合作单位='" & strUnit & "' and 限制项目='" & str星期 & "'"
                   If mrsUnitsReg.RecordCount > 0 Then
                        .TextMatrix(i, j + 1) = Val(Nvl(mrsUnitsReg!数量))
                   End If
                Next
             End If
        Next
    End With
    mrsUnitsReg.Filter = 0
End Sub

Private Sub PicUnit_Resize()
    On Error Resume Next
    lblUnitRegTitle.Move 0, 0, PicUnit.ScaleWidth, lblUnitRegTitle.Height
    Me.lstUnits.Move 0, lblUnitRegTitle.Height, PicUnit.ScaleWidth, PicUnit.ScaleHeight - lblUnitRegTitle.Height
End Sub
Private Sub MoveUnitReg(Optional ByVal str合作单位 As String)
    '对合作单位挂号进行重新分配
    Dim str限制项目  As String, str星期 As String
    Dim lng原限制数量 As Long, lng限制数量 As Long
    Dim lng原分配数量 As Long, lng分配数量  As Long
    Dim bln时段 As Boolean, varCol  As Variant
    Dim str时间段 As String, j As Long, i As Long
    Dim str挂号合作单位     As String
    
    If Not mblnChange Then Exit Sub
    If mbln时段 = False Then Exit Sub
    
    mblnChange = False
    
    On Error GoTo errHandle
    
    varCol = Split(mstr排班, "||")
    If str合作单位 = "" Then
        str挂号合作单位 = lstUnits.Text
    Else
        str挂号合作单位 = str合作单位
    End If
     
    For j = 2 To vsUnits.Rows - 1
        For i = 0 To UBound(varCol)
         
            str时间段 = vsUnits.TextMatrix(j, i * 3 + 0)
            If Trim(vsUnits.TextMatrix(j, i * 3 + 1)) <> "" Then
            lng限制数量 = Val(vsUnits.TextMatrix(j, i * 3 + 1))
            lng分配数量 = Val(vsUnits.TextMatrix(j, i * 3 + 2))
            If Trim(str时间段) = "" Then
                bln时段 = False
            Else
                bln时段 = True
            End If
            str星期 = Split(varCol(i), "|")(0)
            'If str星期 = "周一" Then Stop
            mrsSource.Filter = "限制项目='" & str星期 & "'" & IIf(bln时段, " And 时间段='" & Trim(str时间段) & "'", "")
            mrsUnitsReg.Filter = "限制项目='" & str星期 & "' And 合作单位='" & str挂号合作单位 & "'" & IIf(bln时段, " And 时间段='" & Trim(str时间段) & "'", "")
            
            If mrsSource.RecordCount > 0 Then
                lng原限制数量 = Val(Nvl(mrsSource!数量))
            Else
                lng原限制数量 = 0
            End If
            If mrsUnitsReg.RecordCount > 0 Then
                lng原分配数量 = Val(Nvl(mrsUnitsReg!数量))
            Else
                lng原分配数量 = 0
            End If
            
            lng限制数量 = lng原限制数量 + lng原分配数量 - lng分配数量
            
            If mrsSource.RecordCount > 0 Then
                mrsSource!数量 = lng限制数量
                mrsSource.Update
            Else
                With mrsSource
                    .AddNew
                    !计划Id = mlng计划Id
                    !限制项目 = str星期
                    !序号 = 0
                    !数量 = lng限制数量
                    !时间段 = str时间段
                    .Update
                End With
            End If
            
            If mrsUnitsReg.RecordCount > 0 Then
                With mrsUnitsReg
                    !数量 = lng分配数量
                End With
            Else
                
            With mrsUnitsReg
                .AddNew
                !合作单位 = str挂号合作单位
                !计划Id = mlng计划Id
                !限制项目 = str星期
                !序号 = IIf(mrsSource.RecordCount > 0, mrsSource!序号, 0)
                !数量 = lng分配数量
                !时间段 = str时间段
                .Update
            End With
 
            End If
            End If
            mrsUnitsReg.Filter = 0
            mrsSource.Filter = 0
        Next
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
'------------------------------------------------------------------------
'页面调用过程与方法
'------------------------------------------------------------------------
Public Function InitData() As Boolean
    Dim strSQL As String
    Dim lng计划ID       As Long
    Dim i       As Long
    Dim strTemp As String

    If mlng计划Id = -1 Then Exit Function
    lng计划ID = mlng计划Id

    On Error GoTo Hd

   strSQL = " " & _
        "   Select a.Id as 计划ID,a.计划ID,A.号类,  A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id," & _
        "          A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六,NVL(A.默认时段间隔,5) as 默认时段间隔, " & _
        "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室 " & _
        "   From ( " & vbNewLine & _
        "       Select B.ID,a.id As 计划id, B.号类, A.号码, B.科室id, A.项目id, B.医生姓名, B.医生id, A.周日, A.周一, A.周二, A.周三," & _
        "              A.周四, A.周五, A.周六, B.病案必须, A.分诊方式, A.序号控制, A.生效时间 As 开始时间, A.失效时间 As 终止时间,A.默认时段间隔  As 默认时段间隔 " & _
        "        From 挂号安排 B, 挂号安排计划 A " & _
        "       Where A.安排ID = B.ID And A.Id=[1] " & _
        ") A,收费项目目录 B,部门表 D " & _
        "   Where A.项目id=b.Id(+) And A.科室id =d.Id(+) " & _
        "        "
    Set mrs计划 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng计划Id)
         
    If mrs计划.EOF Then
        ShowMsgbox "未找到指定的号别,请检查!"
        Exit Function
    End If
    
    mbln序号控制 = IIf(Val(Nvl(mrs计划!序号控制)) = 1, True, False)
    mstr排班 = ""

    For i = 0 To 6
        strTemp = Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")

        If Nvl(mrs计划("周" & strTemp)) <> "" Then
            If mstr排班 <> "" Then mstr排班 = mstr排班 & "||"
            mstr排班 = mstr排班 & "周" & strTemp & "|" & Nvl(mrs计划("周" & strTemp))
        End If

    Next
        
    strSQL = "" & _
    "   Select decode(星期,'周日',1,'周一',2,'周二',3,'周三',4,'周四',5,'周五',6,7) as 排序,星期,to_char(开始时间,'HH24')||':00' as 时点,序号,to_char(开始时间,'hh24:mi')||'-' ||to_char(结束时间,'hh24:mi') as 时间范围, " & _
    "               限制数量,是否预约" & _
    "   From  挂号计划时段 " & "   Where 计划ID=[1]" & "   Order by 排序,时点,序号"
    Set mrs时间段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng计划ID)

    If Not mrs时间段.EOF Then mbln时段 = True
    '挂号安排限制
    strSQL = "Select 限制项目,限号数,  限约数,限制项目 as 星期 From  挂号计划限制 where 计划ID=[1]  Order BY 限制项目      "
    Set mrs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng计划Id)
    
    InitData = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function


Private Function InitPage() As Boolean
    Dim i As Long, j As Long
    Dim varCol As Variant, lng限制数量 As Long
    Dim varData As Variant
    On Error GoTo errHandle
    
    If mstr排班 = "" Then Exit Function
    chkDisable.Visible = False
    varCol = Split(mstr排班, "||")
    If mbln时段 Then
        chkDisable.Visible = True
        With vsUnits
                .Clear 1
                .ColWidthMin = 1000
                .Cols = (UBound(varCol) + 1) * 3
                For i = 0 To UBound(varCol)
                    For j = 0 To 2
                          .TextMatrix(0, i * 3 + j) = Split(varCol(i), "|")(0) & "(" & Split(varCol(i), "|")(1) & ")"
                    Next
                    .TextMatrix(1, i * 3 + 0) = "时间段"
                    .TextMatrix(1, i * 3 + 1) = "剩余数量"
                    .TextMatrix(1, i * 3 + 2) = "分配数量"
                Next
            mrs限号.Filter = 0
            For i = 0 To .Cols - 1
                .FixedAlignment(i) = flexAlignCenterCenter
            Next
            .MergeRow(0) = True
            .MergeRow(1) = True
            .AllowUserResizing = flexResizeColumns
            .Editable = flexEDKbdMouse
        End With
        InitPage = True
        Exit Function
    End If
    With vsUnits
        .Cols = UBound(varCol) + 2
        .TextMatrix(0, 0) = "合作单位"
        .TextMatrix(1, 0) = "合作单位"
        .ColWidth(0) = 2000
        For i = 0 To UBound(varCol)
            varData = Split(varCol(i) & "|", "|")
            lng限制数量 = 0
             mrs限号.Filter = "限制项目='" & Split(varCol(i), "|")(0) & "'"
             If mrs限号.RecordCount > 0 Then lng限制数量 = IIf(Val(Nvl(mrs限号!限约数)) = 0, Val(Nvl(mrs限号!限号数)), Val(Nvl(mrs限号!限约数)))
            .TextMatrix(0, i + 1) = varData(0)
            .TextMatrix(1, i + 1) = IIf(varData(1) = "", "无", varData(1)) & "(" & lng限制数量 & ")"
            .Cell(flexcpData, 1, i + 1) = lng限制数量
        Next
        mrs限号.Filter = 0
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "禁用"
        .ColDataType(.Cols - 1) = flexDTBoolean
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .MergeCol(i) = False
        Next
        .ExtendLastCol = False
        .AllowUserResizing = flexResizeColumns
        .MergeRow(0) = True
        '.MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCells = flexMergeRestrictColumns
        .MergeCellsFixed = flexMergeFixedOnly
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Editable = flexEDKbdMouse
        zl_vsGrid_Para_Restore mlngModule, vsUnits, Me.Caption, "三方机构_非时间段"
    End With
    InitPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub chkDisable_Click()
    With chkDisable
        If .Value = 1 Then
            vsUnits.Enabled = False
        Else
            vsUnits.Enabled = True
        End If
        With mrsDisable
            .Filter = "合作单位='" & lstUnits.Text & "'"
            If .RecordCount <> 0 Then
                .MoveFirst
                .Delete adAffectCurrent
                .Update
            End If
            If chkDisable.Value = 1 Then
                .AddNew
                !合作单位 = lstUnits.Text
                .Update
            End If
        End With
    End With
End Sub

Private Function InitRs()
    Dim i         As Long
    Dim j         As Long
    Dim strList() As String
    Dim lng限号数   As Long
    Dim lng限约数   As Long
    Dim rsTmp  As ADODB.Recordset
    Dim strSQL As String
    Dim str限制项目 As String
    Dim str时间段 As String
    Dim lng限制数量 As Long
    Dim lng分配数量 As Long
    
    On Error GoTo errHandle
    
    '初始化 数据集
    With mrsUnitsReg
        Set mrsUnitsReg = New ADODB.Recordset
        mrsUnitsReg.Fields.Append "合作单位", adVarChar, 40
        mrsUnitsReg.Fields.Append "计划ID", adBigInt
        mrsUnitsReg.Fields.Append "限制项目", adVarChar, 10
        mrsUnitsReg.Fields.Append "序号", adBigInt, 18
        mrsUnitsReg.Fields.Append "数量", adBigInt, 18
        mrsUnitsReg.Fields.Append "时间段", adVarChar, 60
        mrsUnitsReg.CursorLocation = adUseClient
        mrsUnitsReg.LockType = adLockOptimistic
        mrsUnitsReg.CursorType = adOpenStatic
        mrsUnitsReg.Open
    End With

    With mrsSource
        Set mrsSource = New ADODB.Recordset
        mrsSource.Fields.Append "计划ID", adBigInt
        mrsSource.Fields.Append "限制项目", adVarChar, 10
        mrsSource.Fields.Append "序号", adBigInt, 18
        mrsSource.Fields.Append "数量", adBigInt, 18
        mrsSource.Fields.Append "时间段", adVarChar, 60
        mrsSource.CursorLocation = adUseClient
        mrsSource.LockType = adLockOptimistic
        mrsSource.CursorType = adOpenStatic
        mrsSource.Open
    End With
    
    With mrsDisable
        Set mrsDisable = New ADODB.Recordset
        mrsDisable.Fields.Append "合作单位", adVarChar, 50
        mrsDisable.CursorLocation = adUseClient
        mrsDisable.LockType = adLockOptimistic
        mrsDisable.CursorType = adOpenStatic
        mrsDisable.Open
    End With
    
    If mstr排班 = "" Then Exit Function
    strList = Split(mstr排班, "||")
    If mbln时段 Then
         '如果是分时段
         
        For i = 0 To UBound(strList)
            mrs时间段.Filter = "星期='" & Split(strList(i), "|")(0) & "' and 是否预约=1"
            If mrs时间段.RecordCount = 0 Then mrs时间段.Filter = "星期='" & Split(strList(i), "|")(0) & "'"
            
            If mrs时间段.RecordCount = 0 Then
               '如果没有设置时间段 不填写时间段
               mrs限号.Filter = "限制项目='" & Split(strList(i), "|")(0) & "'"

               If mrs限号.RecordCount = 0 Then
                   mrs限号.Filter = 0
               Else
                   lng限号数 = Val(Nvl(mrs限号!限号数))
                   lng限约数 = Val(Nvl(mrs限号!限约数))
                   
                   If lng限约数 = 0 Then lng限约数 = lng限号数
                    With mrsSource
                        .AddNew
                        !计划Id = mlng计划Id
                        !限制项目 = Split(strList(i), "|")(0)
                        !序号 = 0
                        !数量 = lng限约数
                        .Update
                    End With
               End If 'mrs限号.recourdcount
               
            Else    'mrs时间段.recordCount=0
                Do While Not mrs时间段.EOF
                    With mrsSource
                        .AddNew
                        !计划Id = mlng计划Id
                        !限制项目 = Split(strList(i), "|")(0)
                        !序号 = Val(Nvl(mrs时间段!序号))
                        !数量 = Val(Nvl(mrs时间段!限制数量))
                        !时间段 = mrs时间段!时间范围
                        .Update
                    End With
                    mrs时间段.MoveNext
                Loop
            End If
        Next
        mrs时间段.Filter = 0
    Else
    
        For i = 0 To UBound(strList)
           '如果没有设置时间段 不填写时间段
            mrs限号.Filter = "限制项目='" & Split(strList(i), "|")(0) & "'"
    
            If mrs限号.RecordCount = 0 Then
                mrs限号.Filter = 0
            Else
                lng限号数 = Val(Nvl(mrs限号!限号数))
                lng限约数 = Val(Nvl(mrs限号!限约数))
                
                If lng限约数 = 0 Then lng限约数 = lng限号数
                '加载初始化数据
                With mrsSource
                    .AddNew
                    !计划Id = mlng计划Id
                    !限制项目 = Split(strList(i), "|")(0)
                    !序号 = 0
                    !数量 = lng限约数
                    .Update
                End With
                
            End If 'mrs限号.recourdcount
        Next
    End If
    
    '已经分配序号
    strSQL = "Select 合作单位, 计划ID, 限制项目, 序号, 数量 From 合作单位计划控制  Where 计划ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng计划Id)

    If rsTmp.RecordCount > 0 Then

        Do While Not rsTmp.EOF
            mrsSource.Filter = "限制项目='" & rsTmp!限制项目 & "' and 序号=" & rsTmp!序号

            With mrsUnitsReg
                .AddNew
                !合作单位 = Nvl(rsTmp!合作单位)
                !计划Id = mlng计划Id
                !限制项目 = Nvl(rsTmp!限制项目)
                !序号 = Val(Nvl(rsTmp!序号))
                !数量 = Val(Nvl(rsTmp!数量))

                If mrsSource.RecordCount > 0 Then
                    !时间段 = mrsSource!时间段
                End If
                
                .Update
            End With

            mrsSource.Filter = 0
            rsTmp.MoveNext
        Loop
    End If
     
    strSQL = "Select Distinct 合作单位 From 合作单位计划控制  Where 计划ID=[1] And 数量 = 0 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng计划Id)
    
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            With mrsDisable
                .AddNew
                !合作单位 = Nvl(rsTmp!合作单位)
                .Update
            End With
            rsTmp.MoveNext
        Loop
    End If
    
     Do While Not mrsSource.EOF
        str限制项目 = mrsSource!限制项目
        str时间段 = mrsSource!时间段
        lng限制数量 = Val(Nvl(mrsSource!数量))
        lng分配数量 = 0
        mrsUnitsReg.Filter = "限制项目='" & str限制项目 & "'" & IIf(Trim(str时间段) <> "", " And 时间段='" & str时间段 & "'", "")
        Do While Not mrsUnitsReg.EOF
            lng分配数量 = Val(Nvl(mrsUnitsReg!数量)) + lng分配数量
            mrsUnitsReg.MoveNext
        Loop
        If lng分配数量 <> 0 Then
           mrsSource!数量 = lng限制数量 - lng分配数量
           mrsSource.Update
        End If
        mrsSource.MoveNext
     Loop
     mrsUnitsReg.Filter = 0
     If mrsUnitsReg.RecordCount > 0 Then mrsUnitsReg.MoveFirst
     mrsSource.Filter = 0
     If mrsSource.RecordCount <> 0 Then mrsSource.MoveFirst
    InitRs = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function chkExists时段(ByVal lng计划ID As Long) As Boolean

    '检查该安排是否具有时段
    Dim strSQL    As String
    Dim rsTmp     As ADODB.Recordset
    Dim blnExists As Boolean
    On Error GoTo Hd
    strSQL = "Select 计划ID From 挂号计划时段 A Where 计划ID=[1] And RowNum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng计划ID)
    chkExists时段 = rsTmp.RecordCount > 0
    Set rsTmp = Nothing
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function InitUntils() As Boolean
    Dim strSQL As String
    Dim i As Long, j        As Long
    Dim lngRow  As Long
    Dim varCol  As Variant
    vsUnits.Clear 1
    vsUnits.ColWidthMin = 1000
    If mstr排班 = "" Then Exit Function
    
    On Error GoTo Hd
    lstUnits.Clear
    
    strSQL = "Select 编码, 名称, 简码, 缺省标志 From 挂号合作单位 Order By 缺省标志 Desc"
    Set mrsUnits = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrsUnits.EOF Then Exit Function
    
    If mbln时段 Then
        Do While Not mrsUnits.EOF
            lstUnits.AddItem Nvl(mrsUnits!名称)
            mrsUnits.MoveNext
        Loop
        If lstUnits.ListCount > 0 Then lstUnits.Selected(0) = True
        InitUntils = True
        Exit Function
    End If
    With vsUnits
        .Clear 1
        .ColWidthMin = 1000
        If mstr排班 = "" Then Exit Function
        .Rows = 2 + mrsUnits.RecordCount
        varCol = Split(mstr排班, "||")
        lngRow = 2
        Do While Not mrsUnits.EOF
            .TextMatrix(lngRow, 0) = mrsUnits!名称
            lngRow = lngRow + 1
            mrsUnits.MoveNext
        Loop
        For i = 2 To .Rows - 1
            mrsDisable.Filter = "合作单位='" & .TextMatrix(i, 0) & "'"
            If mrsDisable.RecordCount <> 0 Then
                .Cell(flexcpChecked, i, .Cols - 1, i, .Cols - 1) = 1
            End If
        Next i
    End With
    InitUntils = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub picUnitReg_Resize()
    On Error Resume Next
    If mbln时段 Then
        lblUnitRegTitle.Move picUnitReg.ScaleLeft, picUnitReg.ScaleTop, picUnitReg.ScaleWidth, lblUnitRegTitle.Height
        chkDisable.Left = picUnitReg.ScaleWidth - chkDisable.Width
        With vsUnits
            .Left = Screen.TwipsPerPixelX * 2
            .Top = lblUnitRegTitle.Top + lblUnitRegTitle.Height + Screen.TwipsPerPixelY * 4
            .Width = picUnitReg.ScaleWidth
            .Height = Me.picUnitReg.ScaleHeight - lblUnitRegTitle.Height - lblUnitRegTitle.Top - Screen.TwipsPerPixelY * 2 - 40 * Screen.TwipsPerPixelY
        End With
   Else
        lblUnitRegTitle.Visible = False
        With vsUnits
            .Left = picUnitReg.ScaleLeft
            .Top = picUnitReg.ScaleTop
            .Width = picUnitReg.ScaleWidth
            .Height = Me.picUnitReg.ScaleHeight
        End With
   End If
  
End Sub

Private Sub vsUnits_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vsUnits.ColIndex("禁用") <> Col Then vsUnits.TextMatrix(Row, Col) = Val(vsUnits.TextMatrix(Row, Col))
     mblnChange = True
End Sub

Public Function SaveData() As Boolean
    Dim i As Long, strSQL As String
    Dim strTmp  As String, cllPro As New Collection
    Dim str合作单位 As String, strPre合作单位 As String, j As Long
    Dim varCol As Variant
    Dim strDisable As String
    If mblnChange Then
        Call MoveUnitReg
    End If
    
    If mbln时段 Then
        For i = 0 To lstUnits.ListCount - 1
            mrsUnitsReg.Filter = "合作单位='" & lstUnits.List(i) & "'"
            strSQL = "Zl_合作单位计划控制_Delete(" & mlng计划Id & ",'" & lstUnits.List(i) & "')"
            zlAddArray cllPro, strSQL
            strDisable = ""
            mrsDisable.Filter = "合作单位='" & lstUnits.List(i) & "'"
            If mrsDisable.RecordCount <> 0 Then
                For j = 1 To vsUnits.Cols - 1
                    If InStr(strDisable, Mid(vsUnits.TextMatrix(0, j), 1, InStr(vsUnits.TextMatrix(0, j), "(") - 1)) = 0 Then
                        If strDisable <> "" Then strDisable = strDisable & "|"
                        strDisable = strDisable & Mid(vsUnits.TextMatrix(0, j), 1, InStr(vsUnits.TextMatrix(0, j), "(") - 1)
                    End If
                Next j
            End If
            If mrsUnitsReg.RecordCount > 0 Then
                With mrsUnitsReg
                    strTmp = ""
                    mrsUnitsReg.Filter = "合作单位='" & lstUnits.List(i) & "' And 数量>0"
                    Do While Not mrsUnitsReg.EOF
                        If strTmp <> "" Then strTmp = strTmp & "|"
                        strTmp = strTmp & !限制项目 & "," & !序号 & "," & !数量
                        mrsUnitsReg.MoveNext
                    Loop
                    If strTmp <> "" And strDisable = "" Then
                        strSQL = "Zl_合作单位计划控制_Insert(" & mlng计划Id & ",'" & lstUnits.List(i) & "','" & strTmp & "')"
                        zlAddArray cllPro, strSQL
                    End If
                End With
                mrsUnitsReg.Filter = 0
            End If
            If strDisable <> "" Then
                strSQL = "Zl_合作单位计划控制_Insert(" & mlng计划Id & ",'" & lstUnits.List(i) & "',Null,Null,'" & strDisable & "')"
                zlAddArray cllPro, strSQL
            End If
        Next
    End If
    If Not mbln时段 Then
        With vsUnits
            For i = 2 To .Rows - 1
               str合作单位 = Trim(.TextMatrix(i, .ColIndex("合作单位")))
               If str合作单位 <> "" Then
                    strSQL = "Zl_合作单位计划控制_Delete(" & mlng计划Id & ",'" & str合作单位 & "')"
                    zlAddArray cllPro, strSQL
                    'Zl_合作单位安排控制_Insert
                    '    安排id_In   合作单位安排控制.安排id%Type,
                    '    合作单位_In 合作单位安排控制.合作单位%Type,
                    '    安排控制_In Varchar2
                    '    --安排控制_in 限制项目,序号1,数量|限制项目,序号2,数量|限制项目,序号3,数量|…………
                    If .Cell(flexcpChecked, i, .Cols - 1, i, .Cols - 1) = 1 Then
                        strTmp = ""
                        For j = 1 To .Cols - 2
                            strTmp = strTmp & "|" & .TextMatrix(0, j)
                        Next
                        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
                        strSQL = "Zl_合作单位计划控制_Insert(" & mlng计划Id & ",'" & str合作单位 & "',Null,Null,'" & strTmp & "')"
                        zlAddArray cllPro, strSQL
                    Else
                        strTmp = ""
                        For j = 1 To .Cols - 1
                            If Val(.TextMatrix(i, j)) <> 0 Then
                                strTmp = strTmp & "|" & .TextMatrix(0, j) & "," & 0 & "," & Val(.TextMatrix(i, j))
                            End If
                        Next
                        If strTmp <> "" Then
                            strTmp = Mid(strTmp, 2)
                            strSQL = "Zl_合作单位计划控制_Insert(" & mlng计划Id & ",'" & str合作单位 & "','" & strTmp & "')"
                            zlAddArray cllPro, strSQL
                        End If
                    End If
               End If
            Next
        End With
    End If
     
    Err = 0: On Error GoTo Errhand:
    mrsDisable.Filter = 0
    strDisable = ""
    Do While Not mrsDisable.EOF
        strDisable = strDisable & "|" & mrsDisable!合作单位
        mrsDisable.MoveNext
    Loop
    If strDisable <> "" Then strDisable = Mid(strDisable, 2)
    zlDatabase.SetPara "禁用合作单位", strDisable, glngSys, 1110
     zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Sub vsUnits_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsUnits.Rows - 1 <= 1 Then Exit Sub
    Call zl_VsGridRowChange(vsUnits, IIf(OldRow = 1, 2, OldRow), NewRow, OldCol, NewCol)
End Sub

Private Sub vsUnits_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
             Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vsUnits_Validate(Cancel As Boolean)
   If mblnChange Then
     MoveUnitReg
   End If
End Sub

Private Sub vsUnits_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim lng分配数量 As Long, lng限制数量 As Long
    Dim lng数量   As Long, str限制项目 As String
    Dim str单位 As String, bln时段 As Boolean
    Dim str时间段 As String, strKey As String
    If Not mbln时段 Then
        With vsUnits
            If .ColIndex("禁用") <> Col Then
                strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
                If .Row < 2 Then Exit Sub
                If zlCommFun.DblIsValid(strKey, 5, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                
                strKey = Format(Abs(Val(strKey)), "####;;;")
                If Val(strKey) > Val(.Cell(flexcpData, 1, Col)) Then
                    MsgBox "数量不能大于限号数(" & Val(.Cell(flexcpData, 1, Col)) & ")", vbOKOnly + vbDefaultButton2 + vbInformation, gstrSysName
                    Call vsUnits_GotFocus
                    Cancel = True: Exit Sub
                End If
                .EditText = strKey
            Else
                With mrsDisable
                    .Filter = "合作单位='" & vsUnits.TextMatrix(Row, 0) & "'"
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        .Delete adAffectCurrent
                        .Update
                    End If
                    If vsUnits.Cell(flexcpChecked, Row, Col, Row, Col) = 2 Then
                        .AddNew
                        !合作单位 = vsUnits.TextMatrix(Row, 0)
                        .Update
                    End If
                End With
            End If
        End With
        Exit Sub
    End If
    
    If Col Mod 3 <> 2 Then Exit Sub
     str时间段 = vsUnits.TextMatrix(Row, Col - 2)
     If Trim(str时间段) = "" Then
         bln时段 = False
     Else
         bln时段 = True
     End If
     str单位 = lstUnits.Text
     mrsSource.Filter = "限制项目='" & str限制项目 & "'" & IIf(bln时段, " And 时间段='" & Trim(str时间段) & "'", "")
     lng分配数量 = Val(vsUnits.EditText)
     str限制项目 = Mid(vsUnits.TextMatrix(0, Col), 1, InStr(vsUnits.TextMatrix(0, Col), "(") - 1)
     mrsSource.Filter = "限制项目='" & str限制项目 & "'" & IIf(bln时段, " And 时间段='" & Trim(str时间段) & "'", "")
     mrsUnitsReg.Filter = "限制项目='" & str限制项目 & "' And 合作单位='" & str单位 & "'" & IIf(bln时段, " And 时间段='" & Trim(str时间段) & "'", "")
     
     If mrsSource.RecordCount = 0 Then
         lng数量 = 0
     Else
         lng数量 = Val(Nvl(mrsSource!数量))
     End If
     If mrsSource.RecordCount = 0 Then mrsSource.Filter = 0: Cancel = True: Exit Sub
     lng数量 = Val(vsUnits.TextMatrix(Row, Col)) + lng数量
     If lng分配数量 > lng数量 Then Cancel = True: Exit Sub
     lng数量 = lng数量 - lng分配数量
     If mrsSource.RecordCount = 0 Then
         With mrsSource
             .AddNew
             !计划Id = mlng计划Id
             !限制项目 = str限制项目
             !序号 = 0
             !数量 = lng数量
             !时间段 = ""
             .Update
         End With
    Else
         With mrsSource
             !数量 = lng数量
             .Update
         End With
    End If
    vsUnits.TextMatrix(Row, Col - 1) = lng数量
    If mrsUnitsReg.RecordCount > 0 Then
             With mrsUnitsReg
                 !数量 = lng分配数量
                 .Update
             End With
     Else
        With mrsUnitsReg
            .AddNew
            !合作单位 = str单位
            !计划Id = mlng计划Id
            !限制项目 = str限制项目
            !序号 = IIf(mrsSource.RecordCount > 0, mrsSource!序号, 0)
            !数量 = lng分配数量
            !时间段 = str时间段
            .Update
        End With
     End If
     mrsUnitsReg.Filter = 0
     mrsSource.Filter = 0
End Sub

Private Sub vsUnits_GotFocus()
    Call zl_VsGridGotFocus(vsUnits)
End Sub

Private Sub vsUnits_LostFocus()
        Call zl_VsGridLOSTFOCUS(vsUnits, GRD_LOSTFOCUS_COLORSEL)
End Sub

Private Sub vsUnits_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    With vsUnits
        If Not mbln时段 Then
            If KeyCode = vbKeyDelete Then
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlVsMoveGridCell vsUnits, 0, vsUnits.Cols - 1
End Sub

Private Sub vsUnits_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
   Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If mbln时段 Then Exit Sub
    With vsUnits
        If Not mbln时段 Then
            If KeyCode = vbKeyDelete Then
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlVsMoveGridCell vsUnits, 0, vsUnits.Cols - 1
End Sub

Private Sub vsUnits_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsUnits
        If mbln时段 Then
            If Col Mod 3 <> 2 Then Cancel = True: Exit Sub
            If .TextMatrix(Row, Col - 1) = "" Then Cancel = True: Exit Sub
             Exit Sub
         End If
         If .Cell(flexcpChecked, Row, .Cols - 1, Row, .Cols - 1) = 1 And Col <> .Cols - 1 Then Cancel = True: Exit Sub
         Select Case Col
         Case .ColIndex("合作单位")
             Cancel = True: Exit Sub
         Case Else
            If .TextMatrix(Row, .ColIndex("合作单位")) = "" Then Cancel = True: Exit Sub
         End Select
    End With
End Sub
Private Sub vsUnits_AfterMoveColumn(ByVal Col As Long, Position As Long)
    If mbln时段 = False Then
        zl_vsGrid_Para_Save mlngModule, vsUnits, Me.Caption, "三方机构_非时间段"
    End If
End Sub
Private Sub vsUnits_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If mbln时段 = False Then
        zl_vsGrid_Para_Save mlngModule, vsUnits, Me.Caption, "三方机构_非时间段"
    End If
End Sub
