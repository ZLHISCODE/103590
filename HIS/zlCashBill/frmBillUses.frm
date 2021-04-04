VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Begin VB.Form frmBillUses 
   Caption         =   "票据明细"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13260
   Icon            =   "frmBillUses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   13260
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraCMD 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   11640
      TabIndex        =   26
      Top             =   540
      Width           =   1455
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   120
         TabIndex        =   16
         Top             =   510
         Width           =   1200
      End
      Begin VB.CommandButton cmdDistant 
         Caption         =   "定位断号(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   18
         Top             =   1710
         Width           =   1200
      End
      Begin VB.TextBox txt号码 
         Height          =   300
         Left            =   150
         TabIndex        =   20
         Top             =   2490
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "定位票据(&F)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   21
         Top             =   2850
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   120
         TabIndex        =   15
         Top             =   30
         Width           =   1200
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   120
         TabIndex        =   17
         Top             =   990
         Width           =   1200
      End
      Begin VB.CommandButton cmdAllDO 
         Caption         =   "全部核对(&A)"
         Height          =   350
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   3510
         Width           =   1200
      End
      Begin VB.CommandButton cmdAllDO 
         Caption         =   "全部取消(&R)"
         Height          =   350
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   3870
         Width           =   1200
      End
      Begin VB.Label lbl号码 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号码(&N)"
         Height          =   180
         Left            =   150
         TabIndex        =   19
         Top             =   2220
         Width           =   630
      End
      Begin VB.Line linBlack 
         BorderColor     =   &H00000000&
         X1              =   120
         X2              =   1300
         Y1              =   1590
         Y2              =   1590
      End
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   90
      ScaleHeight     =   870
      ScaleWidth      =   11475
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11475
      Begin VB.CommandButton cmd领用人 
         Caption         =   "…"
         Height          =   255
         Left            =   5100
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   120
         Width           =   285
      End
      Begin VB.ComboBox cbo使用情况 
         Height          =   300
         Left            =   9270
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   510
         Width           =   1350
      End
      Begin zlIDKind.IDKindNew cboFindType 
         Height          =   300
         Left            =   2310
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   3150
         MaxLength       =   200
         TabIndex        =   4
         Top             =   97
         Width           =   2265
      End
      Begin VB.PictureBox picTimeRange 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   2235
         ScaleHeight     =   390
         ScaleWidth      =   6105
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   6105
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "刷新(&R)"
            Height          =   350
            Left            =   4740
            TabIndex        =   10
            Top             =   30
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker dtpStartDate 
            Height          =   300
            Left            =   0
            TabIndex        =   7
            Top             =   60
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   153288707
            CurrentDate     =   41520
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   300
            Left            =   2550
            TabIndex        =   9
            Top             =   60
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   153288707
            CurrentDate     =   41520
         End
         Begin VB.Label lblEndDate 
            AutoSize        =   -1  'True
            Caption         =   "～"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2265
            TabIndex        =   8
            Top             =   120
            Width           =   210
         End
      End
      Begin VB.ComboBox cbo使用日期 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   525
         Width           =   1350
      End
      Begin VB.ComboBox cbo票种 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   90
         Width           =   1350
      End
      Begin VB.Label lbl使用情况 
         AutoSize        =   -1  'True
         Caption         =   "使用情况"
         Height          =   180
         Left            =   8490
         TabIndex        =   12
         Top             =   570
         Width           =   720
      End
      Begin VB.Line lineSplit 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   10365
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label lbl票种 
         AutoSize        =   -1  'True
         Caption         =   "票据种类"
         Height          =   180
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   720
      End
      Begin VB.Label lbl使用日期 
         AutoSize        =   -1  'True
         Caption         =   "使用时间"
         Height          =   180
         Left            =   90
         TabIndex        =   5
         Top             =   585
         Width           =   720
      End
   End
   Begin VB.TextBox txtInput 
      Height          =   270
      Left            =   7520
      MaxLength       =   200
      TabIndex        =   25
      Top             =   1815
      Visible         =   0   'False
      Width           =   1450
   End
   Begin VB.ComboBox cboResult 
      Height          =   300
      Left            =   6300
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   6735
      Left            =   60
      TabIndex        =   14
      Top             =   930
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColorSel    =   12320767
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorUnpopulated=   -2147483644
      GridColor       =   8421504
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "  号码  |   金额  |    使用时间    |使用人|   使用情况   |     核对时间     |核对人|   核对结果  |      备注     |ID"
      MouseIcon       =   "frmBillUses.frx":0442
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.Label lbl提示 
      AutoSize        =   -1  'True
      Caption         =   "所有人的正常使用明细清单"
      Height          =   180
      Left            =   7800
      TabIndex        =   27
      Top             =   1050
      Width           =   2160
   End
End
Attribute VB_Name = "frmBillUses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mbytInFun As Byte    '0-查看票据使用明细,1-核对票据明细
Private mblnViewCheck As Boolean '当mbytInFun=0时,是否显示核对相关字段
Private mlng票种 As gBillType
Private mblnIsBIll As Boolean '当前票种是否为票据
Private mlng领用ID As Long
Private mdblGiveCount As Double   '该批次票据总张数
Private mstr前缀文本 As String
Private mblnUnClick As Boolean
Private mblnFirst As Boolean

Private Enum Col
    C0号码 = 0
    C_金额 = 1
    C1使用时间 = 2
    C2使用人 = 3
    C3使用情况 = 4
    C4核对时间 = 5
    C5核对人 = 6
    C6核对结果 = 7
    C7备注 = 8
    C8ID = 9
End Enum

Private Sub SetUnChecked(ByVal lngRow As Long)
    With mshDetail
        .TextMatrix(lngRow, Col.C4核对时间) = ""
        .TextMatrix(lngRow, Col.C5核对人) = ""
        .TextMatrix(lngRow, Col.C6核对结果) = ""
        .TextMatrix(lngRow, Col.C7备注) = ""
        
        .RowData(lngRow) = 1  '用于保存时判断
        If Not cmdSave.Enabled Then cmdSave.Enabled = True
    End With
End Sub

Private Sub SetChecked(ByVal lngRow As Long, ByVal lngCol As Long, ByVal strContent As String, Optional ByVal strDate As String)
    With mshDetail
        If lngCol = Col.C6核对结果 Then
            .TextMatrix(lngRow, Col.C4核对时间) = strDate
            .TextMatrix(lngRow, Col.C5核对人) = UserInfo.姓名
            .TextMatrix(lngRow, lngCol) = strContent
        ElseIf lngCol = Col.C7备注 Then
            .TextMatrix(lngRow, lngCol) = strContent
        End If
        
        .RowData(lngRow) = 1 '用于保存时判断
        If Not cmdSave.Enabled Then cmdSave.Enabled = True
    End With
End Sub

Private Sub cboFindType_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnUnClick Then Exit Sub
    cmd领用人.Visible = (cboFindType.IDKind = 1)
    SaveRegInFor g私有模块, Me.Name, "领用批次过滤方式", cboFindType.IDKind
    
    zlControl.ControlSetFocus txtFind: zlControl.TxtSelAll txtFind
End Sub

Private Sub cboResult_LostFocus()
    If cboResult.Visible Then cboResult.Visible = False
End Sub

Private Sub cbo票种_Click()
    Dim lng票种 As gBillType
    Dim strKeyValue As String
    
    If cbo票种.ListIndex <> -1 Then lng票种 = cbo票种.ItemData(cbo票种.ListIndex)
    
    mblnUnClick = True
    If lng票种 = gBillType.就诊卡 Or lng票种 = gBillType.消费卡 Then
        cboFindType.IDKindStr = "领|领用人|0;卡|卡号|0"
    Else
        cboFindType.IDKindStr = "领|领用人|0;发|发票号|0"
    End If
    mblnUnClick = False
    
    GetRegInFor g私有模块, Me.Name, "领用批次过滤方式", strKeyValue
    If Val(strKeyValue) < 1 Or Val(strKeyValue) > 2 Then
        cboFindType.IDKind = 1
    Else
        cboFindType.IDKind = Val(strKeyValue)
    End If
End Sub

Private Sub cbo票种_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub SetRowHiddenAndTipText()
    '设置行显示状态及提示信息
    Dim i As Long, lngCount As Long, dblMoney As Double
    Dim dtDate As Date
    
    On Error GoTo ErrHandler
    With mshDetail
        .Redraw = False
        For i = 1 To .Rows - 1
            If cbo使用日期.Text = "所有" Then
                .RowHeight(i) = .RowHeight(0)
            ElseIf cbo使用日期.Text = "时间范围" Then
                If IsDate(.TextMatrix(i, Col.C1使用时间)) Then
                    dtDate = CDate(.TextMatrix(i, Col.C1使用时间))
                    If dtDate >= dtpStartDate.Value And dtDate <= dtpEndDate.Value Then
                        .RowHeight(i) = .RowHeight(0)
                    Else
                        .RowHeight(i) = 0
                    End If
                Else
                    .RowHeight(i) = 0
                End If
            ElseIf InStr(1, .TextMatrix(i, Col.C1使用时间), cbo使用日期.Text) > 0 Then
                .RowHeight(i) = .RowHeight(0)
            Else
                .RowHeight(i) = 0
            End If
            
            If .RowHeight(i) <> 0 Then
                If Trim(cbo使用情况.Text) <> "" And cbo使用情况.Text <> .TextMatrix(i, Col.C3使用情况) Then
                    .RowHeight(i) = 0
                End If
            End If
            
            If .RowHeight(i) <> 0 Then
                lngCount = lngCount + 1
                dblMoney = dblMoney + Val(.TextMatrix(i, Col.C_金额))
            End If
        Next
        .Redraw = True
    End With
    
    lbl提示.Caption = lbl提示.Tag
    If cbo使用日期.Text <> "所有" Or Trim(cbo使用情况.Text) <> "" Then
        lbl提示.Caption = lbl提示.Caption & IIf(lbl提示.Caption = "", "", ".") & _
            "其中当前选中 " & lngCount & " 张" & IIf(mblnIsBIll, "票据", "卡片")
    End If
    
    If mblnIsBIll Then
        lbl提示.Caption = lbl提示.Caption & IIf(lbl提示.Caption = "", "", ",") & _
            "总金额：" & FormatEx(dblMoney, 2, , , 6)
    End If
    Exit Sub
ErrHandler:
    mshDetail.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo使用情况_Click()
    If mblnUnClick = True Then Exit Sub
    Call SetRowHiddenAndTipText
End Sub

Private Sub cbo使用情况_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo使用日期_Click()
    On Error GoTo errHandle
    If mblnUnClick = True Then Exit Sub
    '问题:29885
    picTimeRange.Visible = False
    If cbo使用日期.Text = "时间范围" Then
        picTimeRange.Visible = True
        If dtpStartDate.Visible And dtpStartDate.Enabled Then dtpStartDate.SetFocus
        Call Form_Resize
        Exit Sub
    End If
    Call SetRowHiddenAndTipText
    Call Form_Resize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo使用日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdAllDO_Click(Index As Integer)
    Dim i As Long, strDate As String
    Dim blnSel As Boolean '是否存在多行选择
    Dim lngRows As Long
    Dim lngStart As Long
    
    With mshDetail
        If .TextMatrix(1, Col.C0号码) = "" Then Exit Sub
        blnSel = .Row <> .RowSel And .RowSel > .Row
        
        If Index = 0 Then
            strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
        lngStart = IIf(blnSel, .Row, 1)
        lngRows = IIf(blnSel, .RowSel, .Rows - 1)
        .Redraw = False
        For i = lngStart To lngRows
            
            If .RowHeight(i) <> 0 Then
                If Index = 0 Then
                   '即使已核对的也重新核对,填写新的核对人和核对时间,不填备注,以前填了的也不用清除
                   Call SetChecked(i, Col.C6核对结果, .TextMatrix(i, Col.C3使用情况), strDate)
                Else
                    '没有核对过的,不必取消核对
                    If Trim(.TextMatrix(i, Col.C6核对结果)) <> "" Then Call SetUnChecked(i)
                End If
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Sub cboResult_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = vbKeyReturn Then
        With mshDetail
            If cboResult.ListIndex <= 0 Then
                Call SetUnChecked(.Row)
            Else
                Call SetChecked(.Row, Col.C6核对结果, Trim(cboResult.Text), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
            End If
            .SetFocus    '调用lostfocus
            .Col = .Col + 1
        End With
    ElseIf KeyAscii >= 32 Then
        If Chr(KeyAscii) > 5 Or Chr(KeyAscii) < 0 Then Exit Sub
        lngIdx = zlControl.CboMatchIndex(cboResult.hWnd, KeyAscii, 0.008)
        If lngIdx = -1 And cboResult.ListCount > 0 And cboResult.ListIndex = -1 Then lngIdx = 0
        cboResult.ListIndex = lngIdx
    End If
End Sub

Private Function SaveData() As Boolean
    Dim i As Long, arrSQL As Variant, blnTrans As Boolean, bytAllChecked As Byte, bytAllCheckOK As Byte
    Dim strDate As String, lngGiveCount As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    With mshDetail
        If .TextMatrix(1, Col.C0号码) = "" Then Exit Function
        arrSQL = Array()
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                strDate = Trim(.TextMatrix(i, Col.C4核对时间))
                If strDate = "" Then
                    strDate = "Null"
                Else
                    strDate = "To_Date('" & .TextMatrix(i, Col.C4核对时间) & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                If mlng票种 = gBillType.消费卡 Then
                    'Zl_消费卡使用记录_Check
                    strSQL = "Zl_消费卡使用记录_Check("
                    '  Id_In       In 消费卡使用记录.Id%Type,
                    strSQL = strSQL & "" & .TextMatrix(i, Col.C8ID) & ","
                    '  核对结果_In In 消费卡使用记录.核对结果%Type,
                    strSQL = strSQL & "" & ZVal(Val(.TextMatrix(i, Col.C6核对结果))) & ","
                    '  核对人_In   In 消费卡使用记录.核对人%Type,
                    strSQL = strSQL & "'" & .TextMatrix(i, Col.C5核对人) & "',"
                    '  备注_In     In 消费卡使用记录.备注%Type,
                    strSQL = strSQL & "'" & .TextMatrix(i, Col.C7备注) & "',"
                    '  核对时间_In In 消费卡使用记录.核对时间%Type
                    strSQL = strSQL & "" & strDate & ")"
                Else
                    'Zl_票据使用明细_Check
                    strSQL = "Zl_票据使用明细_Check("
                    '  Id_In       In 票据使用明细.Id%Type,
                    strSQL = strSQL & "" & .TextMatrix(i, Col.C8ID) & ","
                    '  核对结果_In In 票据使用明细.核对结果%Type,
                    strSQL = strSQL & "" & ZVal(Val(.TextMatrix(i, Col.C6核对结果))) & ","
                    '  核对人_In   In 票据使用明细.核对人%Type,
                    strSQL = strSQL & "'" & .TextMatrix(i, Col.C5核对人) & "',"
                    '  备注_In     In 票据使用明细.备注%Type,
                    strSQL = strSQL & "'" & .TextMatrix(i, Col.C7备注) & "',"
                    '  核对时间_In In 票据使用明细.核对时间%Type
                    strSQL = strSQL & "" & strDate & ")"
                End If
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        Next
    End With
    
    On Error GoTo errH
    If UBound(arrSQL) >= 0 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        
        '检查是否需要填写整体核对记录
        If mlng票种 = gBillType.消费卡 Then
            strSQL = _
                "Select Nvl(Sum(Decode(核对结果, Null, 1, 0)), 0) As 未核对数, Count(Distinct 卡号) As 已使用数" & vbNewLine & _
                "From 消费卡使用记录" & vbNewLine & _
                "Where 领用id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng领用ID)
            If rsTmp!未核对数 = 0 And rsTmp!已使用数 = mdblGiveCount Then
                bytAllChecked = 1
                strSQL = "Select Count(ID) 不相符数 From 消费卡使用记录 Where 领用id = [1] And 核对结果 <> 原因"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng领用ID)
                If rsTmp!不相符数 = 0 Then bytAllCheckOK = 1
            End If
                    
            If bytAllChecked = 1 Then
                strSQL = "Zl_消费卡领用记录_Check(" & mlng领用ID & "," & bytAllCheckOK & ",'" & UserInfo.姓名 & "',Null,1)"
            Else
                '取消整体核对
                strSQL = "Zl_消费卡领用记录_Check(" & mlng领用ID & ",Null,Null,Null,Null)"
            End If
        Else
            strSQL = _
                "Select Nvl(Sum(Decode(核对结果, Null, 1, 0)), 0) As 未核对数, Count(Distinct 号码) As 已使用数" & vbNewLine & _
                "From 票据使用明细" & vbNewLine & _
                "Where 领用id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng领用ID)
            If rsTmp!未核对数 = 0 And rsTmp!已使用数 = mdblGiveCount Then
                bytAllChecked = 1
                strSQL = "Select Count(ID) 不相符数 From 票据使用明细 Where 领用id = [1] And 核对结果 <> 原因"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng领用ID)
                If rsTmp!不相符数 = 0 Then bytAllCheckOK = 1
            End If
                    
            If bytAllChecked = 1 Then
                strSQL = "zl_票据领用记录_check(" & mlng领用ID & "," & bytAllCheckOK & ",'" & UserInfo.姓名 & "',Null,1)"
            Else
                '取消整体核对
                strSQL = "zl_票据领用记录_check(" & mlng领用ID & ",Null,Null,Null,Null)"
            End If
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
       gcnOracle.CommitTrans: blnTrans = False
    End If
    SaveData = True
    
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
      Call SetRowHiddenAndTipText
End Sub

Private Sub cmdSave_Click()
    Dim i As Long
    
    If SaveData Then
        With mshDetail
            For i = 1 To .Rows - 1
                If .RowData(i) = 1 Then .RowData(i) = 0
            Next
        End With
        cmdSave.Enabled = False
    End If
End Sub

Private Sub cmd领用人_Click()
    Dim rsResult As ADODB.Recordset, strSQL As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim bytKind As Byte, varPara As Variant
    
    On Error GoTo ErrHandler
    bytKind = Val(zlStr.NeedCode(cbo票种.Text))
    Call GetOperatorSql(bytKind, strSQL, varPara)
    '坐标定位
    vRect = zlControl.GetControlRect(txtFind.hWnd)
    Set rsResult = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "领用人选择", False, "", "", _
        False, False, True, vRect.Left - 15, vRect.Top, txtFind.Height, blnCancel, False, False, varPara)
    If blnCancel Then zlControl.ControlSetFocus txtFind: zlControl.TxtSelAll txtFind: Exit Sub
    If rsResult Is Nothing Then zlControl.ControlSetFocus txtFind: zlControl.TxtSelAll txtFind: Exit Sub
    If rsResult.RecordCount = 0 Then zlControl.ControlSetFocus txtFind: zlControl.TxtSelAll txtFind: Exit Sub
    
    txtFind.Text = NVL(rsResult!姓名)
    Call txtFind_KeyPress(vbKeyReturn)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetOperatorSql(ByVal bytKind As gBillType, _
    ByRef strSQL As String, ByRef varPara As Variant)
    
    On Error GoTo ErrHandler
    If zlStr.IsHavePrivs(mstrPrivs, "所有操作员") = False Then
        strSQL = _
            "Select A.ID, A.编号, A.姓名" & vbNewLine & _
            "From 人员表 A" & vbNewLine & _
            "Where a.ID=[1]"
        varPara = UserInfo.ID
    Else
        strSQL = _
            "Select Distinct A.ID, A.编号, A.姓名" & vbNewLine & _
            "From 人员表 A, 人员性质说明 B" & vbNewLine & _
            "Where A.ID = B.人员id And B.人员性质 = [1]" & vbNewLine & _
            "      And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            "      And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) Order By 姓名"
        If bytKind > 0 And bytKind <= 7 Then
            '如果是入院登记员，则需要同时设置对应的发卡或预交人员属性这里才显示，病人信息管理同样也有这两项功能了
            varPara = Choose(bytKind, "门诊收费员", "预交收款员", "住院结帐员", "门诊挂号员", _
                "发卡登记人", "发卡登记人", "发卡登记人")
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpEndDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then dtpStartDate.Value = dtpEndDate.Value
End Sub
Private Sub dtpStartDate_Change()
    If dtpStartDate.Value > dtpEndDate.Value Then dtpEndDate.Value = dtpStartDate.Value
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mshDetail.Rows > 1 Then Call SetRow(1)
    If mlng领用ID > 0 Then
        zlControl.ControlSetFocus cbo使用情况
    Else
        zlControl.ControlSetFocus cbo使用日期
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub mshDetail_Click()
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        If .TextMatrix(1, Col.C0号码) = "" Then Exit Sub
        Select Case .Col
            Case Col.C6核对结果
                If .TextMatrix(.Row, .Col) <> "" Then
                    Call zlControl.CboLocate(cboResult, zlCommFun.GetNeedName(.TextMatrix(.Row, .Col)))
                Else
                    Call zlControl.CboLocate(cboResult, zlCommFun.GetNeedName(.TextMatrix(.Row, Col.C3使用情况)))
                End If
                Call SetCboResult
            Case Else
        End Select
    End With
End Sub

Private Sub mshDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
        With mshDetail
            If .TextMatrix(1, Col.C0号码) = "" Then Exit Sub
            Select Case .Col
                Case Col.C6核对结果
                    Call SetUnChecked(.Row)
                Case Col.C7备注
                    Call SetChecked(.Row, Col.C7备注, "")
                Case Else
                
            End Select
        End With
    End If
End Sub

Private Sub mshDetail_KeyPress(KeyAscii As Integer)
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        If .TextMatrix(1, Col.C0号码) = "" Then Exit Sub
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            If .Row = .Rows - 1 And (.Col = Col.C7备注 Or .Col = Col.C6核对结果 And .TextMatrix(.Row, .Col) = "") Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If .Col = Col.C7备注 Then
                    .Row = .Row + 1
                    .Col = Col.C6核对结果
                Else
                    .Col = .Col + 1
                End If
            End If
        Else
            Select Case .Col
                Case Col.C6核对结果
                    Call SetCboResult
                    Call cboResult_KeyPress(KeyAscii)
                Case Col.C7备注
                    If .TextMatrix(.Row, Col.C6核对结果) <> "" Then
                        txtInput.Text = Chr(KeyAscii)
                        txtInput.SelStart = 2
                        Call SetTxtInput
                    End If
                Case Else
                
            End Select
        End If
    End With
End Sub

Private Sub picFilter_Resize()
    On Error Resume Next
    lineSplit.X1 = 0
    lineSplit.X2 = picFilter.ScaleWidth
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strKey As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strKey = Trim(txtFind.Text)
    If strKey = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    Call Select领用批次(txtFind, strKey, Val(zlStr.NeedCode(cbo票种.Text)), cboFindType.IDKind)
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If zlCommFun.ActualLen(txtInput.Text) > txtInput.MaxLength Then
            MsgBox "最多只允许输入" & txtInput.MaxLength & "个字符！", vbInformation, gstrSysName
            Exit Sub
        End If
        If InStr(1, txtInput.Text, "'") > 0 Then
            'MsgBox "注意:单引号是系统禁止输入的特殊字符!", vbInformation, gstrSysName
            Beep
            Beep
            Exit Sub
        End If
        
        With mshDetail
            Call SetChecked(.Row, Col.C7备注, Trim(txtInput.Text))
            txtInput.Visible = False
            .SetFocus  '调用lostfocus
            If .Row = .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                .Row = .Row + 1
                .Col = Col.C6核对结果
            End If
        End With
    End If
End Sub

Private Sub cmdDistant_Click()
    Dim lngRow As Long, bln提醒 As Boolean
    Dim lng前缀 As Long
    
    MousePointer = vbHourglass
    lng前缀 = Len(mstr前缀文本) + 1
    With mshDetail
        lngRow = .Row + 1
        
        While True
            If lngRow > .Rows - 1 Then
                '最后一行
                If bln提醒 = False Then
                    If .Row = 1 Then
                        MsgBox "往下未发现断号情况。", vbInformation, gstrSysName
                        MousePointer = vbDefault
                        Exit Sub
                    Else
                        If MsgBox("往下未发现断号的情况，是否从头开始？", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                            MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    bln提醒 = True
                    lngRow = 1
                Else
                    MsgBox "往下未发现断号情况。", vbInformation, gstrSysName
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            
            If lngRow > 1 Then
                If Val(Mid(.TextMatrix(lngRow - 1, 0), lng前缀)) < Val(Mid(.TextMatrix(lngRow, 0), lng前缀)) - 1 Then
                    '出现断号
                    If .RowHeight(lngRow) = 0 Then
                        If MsgBox("注意:" & vbCrLf & "   已经查找到了断号，但不在当前时间范围内，是否进行定位？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                             If cbo使用日期.Visible Then cbo使用日期.ListIndex = 0:
                        Else
                            MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    Call SetRow(lngRow)
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            lngRow = lngRow + 1
        Wend
     End With
End Sub

Private Sub cmdFind_Click()
'查找指定号码
    Dim strFind As String
    Dim lngRow As Long
    
    If txt号码.Text = "" Then Exit Sub
    If Len(txt号码.Text) > Len(mshDetail.TextMatrix(1, 0)) Then Exit Sub
    
    '把长度补齐
    strFind = UCase(Mid(mshDetail.TextMatrix(1, 0), 1, Len(mshDetail.TextMatrix(1, 0)) - Len(txt号码.Text)) & txt号码.Text)
    With mshDetail
        For lngRow = 1 To mshDetail.Rows - 1
            If mshDetail.TextMatrix(lngRow, 0) = strFind Then
                If .RowHeight(lngRow) = 0 Then
                    If MsgBox("注意:" & vbCrLf & "   你所查找的" & IIf(mblnIsBIll, "号码", "卡号") & "不在当前时间范围内，是否还要进行定位！", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        If cbo使用日期.Visible Then cbo使用日期.ListIndex = 0:
                    Else
                        Exit Sub
                    End If
                End If
                Call SetRow(lngRow)
                Exit Sub
            End If
        Next
    End With
    MsgBox "未找到" & IIf(mblnIsBIll, "号码", "卡号") & "为 " & strFind & " 的使用记录。", vbInformation, gstrSysName
End Sub

Private Sub cmdOK_Click()
    If mbytInFun = 1 Then
        Call SaveData
    End If
    Unload Me
End Sub

Private Sub SetHeader()
    Dim strHead As String, arrTmp As Variant, i As Long
    
    With mshDetail
        .Redraw = False
        If mbytInFun = 0 Then
            .SelectionMode = flexSelectionByRow
        Else
            .SelectionMode = flexSelectionFree
            .BackColorSel = &HE7CFBA
        End If
                
        If mbytInFun = 0 And Not mblnViewCheck Then
            strHead = "号码,1,1000|金额,7,900|使用时间,1,1800|使用人,4,800|使用情况,1,1000"
        Else
            strHead = "号码,1,1000|金额,7,900|使用时间,1,1800|使用人,4,800|使用情况,1,1000|核对时间,1,1800|核对人,4,800|核对结果,1,1000|备注,1,2000|ID,1,0"
        End If
        If mblnIsBIll = False Then strHead = Replace(strHead, "号码", "卡号")
        arrTmp = Split(strHead, "|")
        
        .Cols = UBound(arrTmp) + 1
        For i = 0 To UBound(arrTmp)
            .TextMatrix(0, i) = Split(arrTmp(i), ",")(0)
            .ColAlignment(i) = Split(arrTmp(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(arrTmp(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        If mblnIsBIll = False Then .ColWidth(Col.C_金额) = 0
        .Redraw = True
    End With
End Sub

Private Sub Form_Load()
    mblnFirst = True
    If mbytInFun = 0 And Not mblnViewCheck Then Me.Width = 8000
    
    If mlng票种 = gBillType.就诊卡 Then
        Me.Caption = IIf(mbytInFun = 0, "医疗卡明细清单", "核对医疗卡明细")
    ElseIf mlng票种 = gBillType.消费卡 Then
        Me.Caption = IIf(mbytInFun = 0, "消费卡明细清单", "核对消费卡明细")
    Else
        Me.Caption = IIf(mbytInFun = 0, "票据明细清单", "核对票据明细")
    End If
    Call InitContext
    If mlng票种 > 0 And mlng票种 - 1 < cbo票种.ListCount Then cbo票种.ListIndex = mlng票种 - 1
    
    Call RestoreFlexState(mshDetail, Me.Caption)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(mshDetail, Me.Caption)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If mbytInFun = 0 Then
        Set lbl提示.Container = Me
        lbl提示.Top = 100: lbl提示.Left = 100
        mshDetail.Top = lbl提示.Height + lbl提示.Top + 20
    Else
        Set lbl提示.Container = picFilter
        If picTimeRange.Visible Then
            lbl使用情况.Left = lbl使用日期.Left
            cbo使用情况.Top = cbo使用日期.Top + cbo使用日期.Height + 50
        Else
            lbl使用情况.Left = cbo使用日期.Left + cbo使用日期.Width + 220
            cbo使用情况.Top = cbo使用日期.Top
        End If
        lbl使用情况.Top = cbo使用情况.Top + (cbo使用情况.Height - lbl提示.Height) / 2
        cbo使用情况.Left = lbl使用情况.Left + lbl使用情况.Width + 50
        lbl提示.Left = cbo使用情况.Left + cbo使用情况.Width + 220
        lbl提示.Top = cbo使用情况.Top + (cbo使用情况.Height - lbl提示.Height) / 2
        picFilter.Height = cbo使用情况.Top + cbo使用情况.Height
        mshDetail.Top = picFilter.Height + picFilter.Top + 20
    End If
    
    mshDetail.Height = Me.ScaleHeight - mshDetail.Top - 120
    If Me.ScaleWidth > 3000 Then
        fraCMD.Left = Me.ScaleWidth - fraCMD.Width - 120
        mshDetail.Width = fraCMD.Left - mshDetail.Left - 120
    End If
    
    picFilter.Left = 0: picFilter.Width = Me.ScaleWidth
End Sub

Public Sub ShowMe(ByVal frmOwner As Form, ByVal strPrivs As String, _
    ByVal bytInFun As Byte, ByVal blnViewCheck As Boolean, ByVal blnNOMoved As Boolean, _
    ByVal lng票种 As gBillType, ByVal lng领用ID As Long, ByVal str前缀 As String, _
    Optional strCondition As String, Optional lng原因 As Long, Optional lng性质 As Long, _
    Optional str使用人 As String, Optional str提示 As String)
    '参数:bytInFun:0-查看票据明细,1-核对票据明细
    '   blnViewCheck:当bytInFun=0时,是否显示核对相关字段
    Dim strResult As String, arrTmp As Variant
    Dim i As Integer
    
    On Error GoTo errH
    mstrPrivs = strPrivs
    mbytInFun = bytInFun
    mblnViewCheck = blnViewCheck
    mlng票种 = lng票种
    mlng领用ID = lng领用ID
    mstr前缀文本 = str前缀
    mblnIsBIll = CurrentIsBill(mlng票种)
    lbl号码.Caption = IIf(mblnIsBIll, "号码(&N)", "卡号(&N)")
    cmdFind.Caption = IIf(mblnIsBIll, "定位票据(&F)", "定位卡号(&F)")
    
    If RefrashData(lng领用ID, blnNOMoved, strCondition, lng原因, lng性质, str使用人, str提示) = False Then Exit Sub
    
    cboResult.Visible = False
    txtInput.Visible = False
    If mbytInFun = 0 Then
        cmdOK.Caption = "退出(&X)"
        cmdOK.Cancel = True
        cmdCancel.Visible = False
        cmdSave.Visible = False
        cmdAllDO(0).Visible = False
        cmdAllDO(1).Visible = False
        picFilter.Visible = False
        lbl提示.Left = picFilter.Left
    Else
        If mlng票种 = gBillType.消费卡 Then
            strResult = " ,1-正常使用,2-作废收回,3-换卡发出,4-换卡收回,5-报损"
            Call zlControl.CboSetWidth(cboResult.hWnd, 1500)
        Else
            strResult = " ,1-正常使用,2-作废收回,3-重打发出,4-重打收回,5-报损,6-红票发出"
            Call zlControl.CboSetWidth(cboResult.hWnd, 800)
        End If
        arrTmp = Split(strResult, ",")
        For i = 0 To UBound(arrTmp)
            cboResult.AddItem arrTmp(i)
            cbo使用情况.AddItem arrTmp(i)
        Next
        picFilter.Visible = True
    End If
    frmBillUses.Show vbModal, frmOwner
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Function RefrashData(ByVal lng领用ID As Long, Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal strCondition As String, Optional ByVal lng原因 As Long, _
    Optional ByVal lng性质 As Long, Optional ByVal str使用人 As String, Optional ByVal str提示 As String) As Boolean
    '加载数据
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, j As Long
    Dim strTemp As String, str使用日期 As String
    Dim strMinDate As String, strMaxDate As String
    Dim varData As Variant, strNOs As String
    
    On Error GoTo errHandle
    If mlng票种 = gBillType.消费卡 Then
        strSQL = _
            "Select 卡号 As 号码, '' As 票据金额, To_Char(使用时间, 'yyyy-mm-dd hh24:mi:ss') As 使用时间, 使用人," & vbNewLine & _
            "       Decode(原因, 1, '1-正常使用', 2, '2-作废收回', 3, '3-换卡发出', 4, '4-换卡收回', '5-报损') As 使用情况," & vbNewLine & _
            "       To_Char(核对时间, 'yyyy-mm-dd hh24:mi:ss') As 核对时间, 核对人," & vbNewLine & _
            "       Decode(核对结果, 1, '1-正常使用', 2, '2-作废收回', 3, '3-重打发出', 4, '4-重打收回', 5,'5-报损','') as 核对结果, 备注, ID" & vbNewLine & _
            "From " & IIf(blnNOMoved, "H", "") & "消费卡使用记录" & vbNewLine & _
            "Where 领用id = [1] " & strCondition & vbNewLine & _
            "Order By 卡号,使用时间"
        If mbytInFun = 0 And Not mblnViewCheck Then
            strSQL = _
                "Select 卡号 As 号码, '' As 票据金额, To_Char(使用时间, 'yyyy-mm-dd hh24:mi:ss') As 使用时间, 使用人," & vbNewLine & _
                "       Decode(原因, 1, '1-正常使用', 2, '2-作废收回', 3, '3-换卡发出', 4, '4-换卡收回', '5-报损') As 使用情况" & vbNewLine & _
                "From " & IIf(blnNOMoved, "H", "") & "消费卡使用记录" & vbNewLine & _
                "Where 领用id = [1] " & strCondition & vbNewLine & _
                "Order By 卡号,使用时间"
        End If
    Else
        strSQL = _
            "Select 号码, Trim(To_Char(票据金额, '99999999990.00')) As 票据金额, To_Char(使用时间, 'yyyy-mm-dd hh24:mi:ss') As 使用时间, 使用人," & vbNewLine & _
            "       Decode(原因, 1, '1-正常使用', 2, '2-作废收回', 3, '3-重打发出', 4, '4-重打收回', 6, '6-红票发出', '5-报损') As 使用情况," & vbNewLine & _
            "       To_Char(核对时间, 'yyyy-mm-dd hh24:mi:ss') As 核对时间, 核对人," & vbNewLine & _
            "       Decode(核对结果, 1, '1-正常使用', 2, '2-作废收回', 3, '3-重打发出', 4, '4-重打收回', 5,'5-报损','') as 核对结果, 备注, ID" & vbNewLine & _
            "From " & IIf(blnNOMoved, "H", "") & "票据使用明细" & vbNewLine & _
            "Where 领用id = [1] " & strCondition & vbNewLine & _
            "Order By 号码,使用时间"
        If mbytInFun = 0 And Not mblnViewCheck Then
            strSQL = _
                "Select 号码, Trim(To_Char(票据金额, '99999999990.00')) As 票据金额, To_Char(使用时间, 'yyyy-mm-dd hh24:mi:ss') As 使用时间, 使用人," & vbNewLine & _
                "       Decode(原因, 1, '1-正常使用', 2, '2-作废收回', 3, '3-重打发出', 4, '4-重打收回', 6, '6-红票发出', '5-报损') As 使用情况" & vbNewLine & _
                "From " & IIf(blnNOMoved, "H", "") & "票据使用明细" & vbNewLine & _
                "Where 领用id = [1] " & strCondition & vbNewLine & _
                "Order By 号码,使用时间"
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取数据", lng领用ID, lng原因, lng性质, str使用人)
    If rsTmp.RecordCount = 0 Then
        mshDetail.Clear
        mshDetail.Rows = 2
    Else
        Set mshDetail.DataSource = rsTmp
    End If
    Call SetHeader
    
    lbl提示.Tag = str提示 & IIf(str提示 = "", "", ",") & "共计 " & rsTmp.RecordCount & " 张" & IIf(mblnIsBIll, "票据", "卡片")
    str使用日期 = ""
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            strTemp = "|" & Format(rsTmp!使用时间, "yyyy-MM-DD")
            If InStr(1, str使用日期 & "|", strTemp & "|") = 0 Then str使用日期 = str使用日期 & strTemp
            rsTmp.MoveNext
        Next
    End If
    If str使用日期 <> "" Then str使用日期 = Mid(str使用日期, 2)
    
    mblnUnClick = True
    If cbo使用情况.ListCount > 0 Then cbo使用情况.ListIndex = 0
    '按日期重小到大排序
    cbo使用日期.Clear
    cbo使用日期.AddItem "所有": cbo使用日期.ListIndex = cbo使用日期.NewIndex
    cbo使用日期.AddItem "时间范围"
    mblnUnClick = False
    
    varData = Split(str使用日期, "|")
    For i = 0 To UBound(varData)
        If varData(i) <> "" Then
            For j = i + 1 To UBound(varData)
                If varData(j) < varData(i) Then
                    strTemp = varData(i)
                    varData(i) = varData(j)
                    varData(j) = strTemp
                End If
            Next
            If varData(i) < strMinDate Or strMinDate = "" Then strMinDate = varData(i)
            If varData(i) > strMaxDate Then strMaxDate = varData(i)
            cbo使用日期.AddItem varData(i)
        End If
    Next
    
    dtpStartDate.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    dtpEndDate.MaxDate = dtpStartDate.MaxDate
    If strMinDate <> "" And IsDate(strMinDate) Then
        dtpStartDate.MinDate = Format(CDate(strMinDate), "yyyy-mm-dd 00:00:00")
        dtpStartDate.Value = dtpStartDate.MinDate
        dtpEndDate.MinDate = dtpStartDate.MinDate
        If IsDate(strMaxDate) Then
            dtpEndDate.Value = Format(CDate(strMaxDate), "yyyy-mm-dd 23:59:59")
        Else
            dtpEndDate.Value = Format(dtpStartDate.MinDate, "yyyy-mm-dd 23:59:59")
        End If
    End If
    
    If mbytInFun <> 0 Then
        mdblGiveCount = 0
        If mlng票种 = gBillType.消费卡 Then
            strSQL = _
                "Select To_Number(Replace(终止卡号, 前缀文本)) - To_Number(Replace(开始卡号, 前缀文本))+1 As 张数," & vbNewLine & _
                "       前缀文本" & vbNewLine & _
                "From 消费卡领用记录" & vbNewLine & _
                "Where ID = [1]"
        Else
            strSQL = _
                "Select To_Number(Replace(终止号码, 前缀文本)) - To_Number(Replace(开始号码, 前缀文本))+1 As 张数," & vbNewLine & _
                "       前缀文本" & vbNewLine & _
                "From 票据领用记录" & vbNewLine & _
                "Where ID = [1]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng领用ID)
        If rsTmp.RecordCount > 0 Then
            mdblGiveCount = rsTmp!张数
            mstr前缀文本 = NVL(rsTmp!前缀文本)
        End If
    End If
    picTimeRange.Visible = False: Call Form_Resize
    Call SetRowHiddenAndTipText
    
    RefrashData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mshDetail_DblClick()
    Dim strReportNO As String, strInvoiceNO As String
    
    If mlng票种 = gBillType.消费卡 Then Exit Sub
    With mshDetail
        If .TextMatrix(1, Col.C0号码) = "" Then Exit Sub
        Select Case .Col
            Case Col.C7备注
                If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
                If .TextMatrix(.Row, Col.C6核对结果) = "" Then Exit Sub
                
                Call SetTxtInput
                txtInput.Text = .TextMatrix(.Row, .Col)
                Call zlControl.TxtSelAll(txtInput)
            Case Else
                strReportNO = "ZL" & glngSys \ 100 & "_INSIDE_1501"
                strInvoiceNO = .TextMatrix(.Row, Col.C0号码)
                Call ReportOpen(gcnOracle, glngSys, strReportNO, Me, "票据号=" & strInvoiceNO & "", "票种=" & mlng票种, "ReportFormat=" & mlng票种, 1)
        End Select
    End With
End Sub

Private Sub SetCboResult()
    With mshDetail
        cboResult.Left = .Left + .CellLeft - 15
        cboResult.Top = .Top + .CellTop - 15
        cboResult.Width = .CellWidth + 15
        cboResult.Visible = True
        cboResult.SetFocus
    End With
End Sub

Private Sub SetTxtInput()
    With mshDetail
        txtInput.Left = .Left + .CellLeft - 15
        txtInput.Top = .Top + .CellTop - 15
        txtInput.Width = .CellWidth + 15
        txtInput.Height = .CellHeight
        txtInput.Visible = True
        txtInput.SetFocus
    End With
End Sub

Private Sub mshDetail_LeaveCell()
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        If .TextMatrix(1, Col.C0号码) = "" Then Exit Sub
        If cboResult.Visible Then
            If .TextMatrix(.Row, .Col) <> Trim(cboResult.Text) Then
                If cboResult.ListIndex <= 0 Then
                    Call SetUnChecked(.Row)
                Else
                    Call SetChecked(.Row, Col.C6核对结果, Trim(cboResult.Text), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
                End If
            End If
        ElseIf txtInput.Visible Then
            If .TextMatrix(.Row, .Col) <> Trim(txtInput.Text) Then
                If zlCommFun.ActualLen(txtInput.Text) > txtInput.MaxLength Then
                    MsgBox "最多只允许输入" & txtInput.MaxLength & "个字符!", vbInformation, gstrSysName
                    Exit Sub
                End If
                If InStr(1, txtInput.Text, "'") > 0 Then
                    'MsgBox "注意:单引号是系统禁止输入的特殊字符!", vbInformation, gstrSysName
                    Beep
                    Beep
                    Exit Sub
                End If
                Call SetChecked(.Row, Col.C7备注, Trim(txtInput.Text))
            End If
        End If
    End With
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngCol As Long
    
    With mshDetail
        If .TextMatrix(1, Col.C0号码) = "" Then Exit Sub
        If Button = 1 And .MousePointer = 99 Then
            lngCol = .MouseCol
            If .TextMatrix(0, lngCol) = "" Then Exit Sub
            
            .ColData(lngCol) = (.ColData(lngCol) + 1) Mod 2
            
            .Redraw = False
            .Col = lngCol: .ColSel = lngCol   '排序依据
            .Sort = IIf(.ColData(lngCol) = 1, 6, 5)
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
        End If
    End With
End Sub


Private Sub txtInput_LostFocus()
    If txtInput.Visible Then txtInput.Visible = False
End Sub

Private Sub txt号码_GotFocus()
    Call zlControl.TxtSelAll(txt号码)
End Sub

Private Sub txt号码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click
        zlControl.TxtSelAll txt号码
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub SetRow(ByVal lngRow As Long)
    Dim lngTop As Long
    With mshDetail
        .Row = lngRow
        lngTop = lngRow - 1
        If lngTop < 1 Then lngTop = 1
        If .RowIsVisible(lngTop) = False Then
            .TopRow = lngTop
        End If
        If mbytInFun = 0 Then
            .Col = 0
            .ColSel = .Cols - 1
        Else
            .Col = Col.C6核对结果
        End If
    End With
End Sub

Private Sub InitContext()
    Dim bln药店 As Boolean
    
    On Error GoTo errHandle
    bln药店 = (glngSys \ 100 = 8)
    
    cbo票种.Clear
    If bln药店 Then
        cbo票种.AddItem "1-收费收据":        cbo票种.ItemData(cbo票种.NewIndex) = 1
        cbo票种.AddItem "5-会员卡":          cbo票种.ItemData(cbo票种.NewIndex) = 5
    Else
        If InStr(1, mstrPrivs, ";收费收据;") > 0 Then
            cbo票种.AddItem "1-收费收据":        cbo票种.ItemData(cbo票种.NewIndex) = 1
        End If
        If InStr(1, mstrPrivs, ";预交收据;") > 0 Or _
          (InStr(1, mstrPrivs, ";预交门诊票据;") > 0 _
          Or InStr(1, mstrPrivs, ";预交住院票据;") > 0) Then
            cbo票种.AddItem "2-预交收据":        cbo票种.ItemData(cbo票种.NewIndex) = 2
        End If
        If InStr(1, mstrPrivs, ";结帐收据;") > 0 Then
          cbo票种.AddItem "3-结帐收据":        cbo票种.ItemData(cbo票种.NewIndex) = 3
        End If
        If InStr(1, mstrPrivs, ";挂号收据;") > 0 Then
          cbo票种.AddItem "4-挂号收据":        cbo票种.ItemData(cbo票种.NewIndex) = 4
        End If
        If InStr(1, mstrPrivs, ";医疗卡;") > 0 Then
           cbo票种.AddItem "5-医疗卡":          cbo票种.ItemData(cbo票种.NewIndex) = 5
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "消费卡") Then
           cbo票种.AddItem "6-消费卡": cbo票种.ItemData(cbo票种.NewIndex) = 6
        End If
'        cbo票种.AddItem "1-收费收据":        cbo票种.ItemData(cbo票种.NewIndex) = 1
'        cbo票种.AddItem "2-预交收据":        cbo票种.ItemData(cbo票种.NewIndex) = 2
'        cbo票种.AddItem "3-结帐收据":        cbo票种.ItemData(cbo票种.NewIndex) = 3
'        cbo票种.AddItem "4-挂号收据":        cbo票种.ItemData(cbo票种.NewIndex) = 4
'        cbo票种.AddItem "5-医疗卡":          cbo票种.ItemData(cbo票种.NewIndex) = 5
'        cbo票种.AddItem "6-消费卡":          cbo票种.ItemData(cbo票种.NewIndex) = 6
        cbo票种.ListIndex = 0
    End If
    
    cboFindType.NotAutoAppendKind = True
    cboFindType.ShowSortName = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Select领用批次(ByVal objCtl As Object, _
    ByVal strKey As String, ByVal int票种 As gBillType, ByVal bytMode As Byte) As Boolean
    '功能:选择指定的领用批次
    '入参:
    '     strKey-输入的建值
    '     int票种-当前选择的票种
    '     bytMode-查找模式：1-按领用人查找，2-按发票号查找
    '出参:
    '返回:查找成功,返回true,否则返回False
    Dim rsTemp As ADODB.Recordset, strWhere As String, strSQL As String
    Dim blnCancel As Boolean, vRect As RECT, blnFind As Boolean
    Dim str使用类别 As String
    
    Err = 0: On Error GoTo ErrHand:
    If strKey = "" Then Exit Function
    
    If bytMode = 1 Then '按领用人查找
        If IsNumeric(strKey) Then '1.输入全是数字时只匹配编码
            strWhere = " And 编号 Like [2]"
        ElseIf zlCommFun.IsCharAlpha(strKey) Then '2.输入全是字母时只匹配简码
            strWhere = " And 简码 Like [2]"
        ElseIf zlStr.IsCharChinese(strKey) Then '2.输入含有汉字时只匹配姓名
            strWhere = " And 姓名 Like [2]"
        Else
            strWhere = " And (姓名 Like [2] Or 简码 Like [2] Or 编号 Like [2])"
        End If
        strKey = GetMatchingSting(strKey, False)
        
        strWhere = " And a.领用人 In (Select 姓名 From  人员表 Where 1=1 " & strWhere & ")"
    Else '按发票号查找
        '说明： And 号码 Like '%%' 这一句的目的是使得记录只有一行时不弹出选择器
        If int票种 = gBillType.消费卡 Then
            strWhere = " And a.Id In(Select 领用ID From 消费卡使用记录 Where 卡号 = [2] And 卡号 Like '%%')"
        Else
            strWhere = " And a.Id In(Select 领用ID From 票据使用明细 Where 票种 = [1] And 号码 = [2] And 号码 Like '%%')"
        End If
    End If
    
    If int票种 = gBillType.就诊卡 Then
        strSQL = _
            "Select a.Id, Nvl(b.名称, '就诊卡') As 使用类别, a.开始号码, a.终止号码, a.领用人," & vbNewLine & _
            "       Decode(a.使用方式, 1, '自用', '共用') As 使用方式, a.备注, a.登记人, To_Char(a.登记时间, 'yyyy-mm-dd') As 领用时间," & vbNewLine & _
            "       Row_Number() Over(Partition By a.领用人 Order By a.登记时间 Desc) As 组内行号" & vbNewLine & _
            "From 票据领用记录 A, 医疗卡类别 B" & vbNewLine & _
            "Where To_Number(Nvl(a.使用类别, '0')) = b.Id(+) And a.票种 = [1]" & strWhere
    ElseIf int票种 = gBillType.消费卡 Then
        strSQL = _
            "Select a.Id, Nvl(b.名称, '消费卡') As 使用类别, a.开始卡号 As 开始号码, a.终止卡号 As 终止号码, a.领用人," & vbNewLine & _
            "       Decode(a.使用方式, 1, '自用', '共用') As 使用方式, a.备注, a.登记人, To_Char(a.登记时间, 'yyyy-mm-dd') As 领用时间," & vbNewLine & _
            "       Row_Number() Over(Partition By a.领用人 Order By a.登记时间 Desc) As 组内行号" & vbNewLine & _
            "From 消费卡领用记录 A, 消费卡类别目录 B" & vbNewLine & _
            "Where To_Number(Nvl(a.接口编号, '0')) = b.编号(+)" & strWhere
    Else
        If int票种 = gBillType.收费收据 Or int票种 = gBillType.结帐收据 Then
            str使用类别 = "a.使用类别,"
        ElseIf int票种 = gBillType.预交收据 Then
            str使用类别 = "Decode(Nvl(a.使用类别,'0'),'0','','1','门诊','住院') As 使用类别,"
        End If
        
        strSQL = _
            "Select a.Id, " & str使用类别 & " a.开始号码, a.终止号码, a.领用人, Decode(a.使用方式, 1, '自用', '共用') As 使用方式," & vbNewLine & _
            "       a.备注, a.登记人, To_Char(a.登记时间, 'yyyy-mm-dd') As 领用时间," & vbNewLine & _
            "       Row_Number() Over(Partition By a.领用人 Order By a.登记时间 Desc) As 组内行号" & vbNewLine & _
            "From 票据领用记录 A" & vbNewLine & _
            "Where a.票种 = [1] " & strWhere
    End If
    
    strSQL = _
        "Select Id," & IIf(int票种 = gBillType.挂号收据, "", " 使用类别,") & _
        "       开始号码, 终止号码, 领用人, 使用方式, 备注, 登记人, 领用时间" & vbNewLine & _
        "From (" & strSQL & ")" & vbNewLine
    If bytMode = 1 Then '按领用人查找时，显示最近10次领用记录
        strSQL = strSQL & _
            "Where 组内行号 < 11"
    End If
    
    '坐标定位
    vRect = zlControl.GetControlRect(objCtl.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "领用记录选择", False, "", "", False, False, True, _
        vRect.Left - 15, vRect.Top, objCtl.Height, blnCancel, False, False, int票种, UCase(strKey))
   
   If blnCancel Then
        zlControl.ControlSetFocus objCtl: zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "未找到满足条件的领用记录，请检查！"
        zlControl.ControlSetFocus objCtl: zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "未找到满足条件的领用记录，请检查！"
        zlControl.ControlSetFocus objCtl: zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    mlng票种 = int票种
    mblnIsBIll = CurrentIsBill(mlng票种)
    mlng领用ID = Val(NVL(rsTemp!ID))
    lbl号码.Caption = IIf(mblnIsBIll, "号码(&N)", "卡号(&N)")
    cmdFind.Caption = IIf(mblnIsBIll, "定位票据(&F)", "定位卡号(&F)")
    If bytMode = 1 Then txtFind.Text = NVL(rsTemp!领用人)
    
    Call RefrashData(mlng领用ID)
    zlCommFun.PressKey vbKeyTab
    
    Select领用批次 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

