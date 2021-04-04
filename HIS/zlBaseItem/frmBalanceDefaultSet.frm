VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBalanceDefaultSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "缺省结算方式设置"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5610
   Icon            =   "frmBalanceDefaultSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4290
      TabIndex        =   4
      Top             =   3240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4290
      TabIndex        =   3
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4290
      TabIndex        =   2
      Top             =   180
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf付款方式 
      Height          =   2985
      Left            =   120
      TabIndex        =   1
      Top             =   630
      Width           =   3975
      _cx             =   7011
      _cy             =   5265
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
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   260
      RowHeightMax    =   260
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceDefaultSet.frx":000C
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
   Begin VB.Label lbl提示 
      Caption         =   "    请设置各医疗付款方式在该场合的缺省结算方式。"
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4110
   End
End
Attribute VB_Name = "frmBalanceDefaultSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr场合 As String
Dim mblnItem As Boolean
Dim mintSuccess As Integer
Dim mblnChange As Boolean     '是否改变了

Public Function ShowMe(frmParent As Object, ByVal str场合 As String) As Boolean
'功能:用来与调用的结算方式管理窗口进行通讯的程序
'参数:str场合     当前编辑的结算方式的编码
'返回值:编辑成功返回True,否则为False
    On Error GoTo ErrHandler
    
    mstr场合 = str场合
    
    Me.Show vbModal, frmParent
    ShowMe = mintSuccess > 0
    
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, 5
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler
    If SaveData() = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
    mblnChange = False
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveData() As Boolean
    '功能:保存编辑的内容到结算方式表中
    '参数:
    '返回值:成功返回True,否则为False
    Dim i As Integer, str缺省 As String
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    '把所有缺省结算方式做成一个串
    '格式：医疗付款方式1:结算方式1;医疗付款方式2:结算方式2;...
    With vsf付款方式
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("缺省结算方式"))) <> "" Then
                str缺省 = str缺省 & ";"
                str缺省 = str缺省 & Trim(.TextMatrix(i, .ColIndex("医疗付款方式"))) & ":"
                str缺省 = str缺省 & Trim(.TextMatrix(i, .ColIndex("缺省结算方式")))
            End If
        Next
    End With
    If str缺省 <> "" Then str缺省 = Mid(str缺省, 2)
    
    '修改
    strSQL = "zl_结算方式应用_update( '" & mstr场合 & "','',1,'" & str缺省 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Dim strSQL As String
    Dim rs结算方式 As New ADODB.Recordset
    Dim rs付款方式 As New ADODB.Recordset, lngRow As Long
    Dim str结算方式 As String
    
    mblnChange = False
    mintSuccess = 0
    
    lbl提示.Caption = Replace(lbl提示.Caption, "该场合", "『" & mstr场合 & "』结算场合")
    
    '可选择置的结算方式：性质为1,2,7,8，不是应付款的
    strSQL = "Select B.结算方式" & _
            " From 结算方式 A,结算方式应用 B" & _
            " Where A.名称=B.结算方式 And B.应用场合=[1] And b.付款方式 Is Null" & _
            "       And A.性质 In(1,2,7,8) And nvl(A.应付款,0)<>1" & _
            " Order by A.编码"
    Set rs结算方式 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr场合)
    If rs结算方式.EOF Then
        MsgBox "    『" & mstr场合 & "』结算场合没有可用于设置为缺省结算方式的结算方式，" & _
            "请先对『" & mstr场合 & "』结算场合设置性质为(1,2,7,8)，并且不是应付款的结算方式！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    Do Until rs结算方式.EOF
        str结算方式 = str结算方式 & "|" & Nvl(rs结算方式!结算方式)
        rs结算方式.MoveNext
    Loop
    vsf付款方式.ColComboList(vsf付款方式.ColIndex("缺省结算方式")) = " " & str结算方式
    
    '医疗付款方式，以及已设置的缺省结算方式
    strSQL = "Select a.名称 As 付款方式, b.结算方式" & vbNewLine & _
            " From 医疗付款方式 A, 结算方式应用 B" & vbNewLine & _
            " Where a.名称 = b.付款方式(+) And b.应用场合(+) = [1]" & vbNewLine & _
            "       And b.付款方式(+) Is Not Null" & vbNewLine & _
            " Order By a.编码"
    Set rs付款方式 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr场合)
    
    With vsf付款方式
        .Clear 1
        .Editable = flexEDKbdMouse
        
        .Rows = rs付款方式.RecordCount + 1
        lngRow = 1
        Do Until rs付款方式.EOF
            .TextMatrix(lngRow, .ColIndex("医疗付款方式")) = Nvl(rs付款方式!付款方式)
            .TextMatrix(lngRow, .ColIndex("缺省结算方式")) = Nvl(rs付款方式!结算方式)
            lngRow = lngRow + 1
            rs付款方式.MoveNext
        Loop
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
End Sub

Private Sub vsf付款方式_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsf付款方式_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsf付款方式.ColIndex("医疗付款方式") Then Cancel = True
End Sub

Private Sub vsf付款方式_GotFocus()
    With vsf付款方式
        If .Row < .FixedRows And .Rows > .FixedRows Then .Row = .FixedRows
    End With
End Sub


