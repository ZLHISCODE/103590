VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClinicDoctorTitleSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "职称标识设置"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5610
   Icon            =   "frmClinicDoctorTitleSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   360
      Left            =   4320
      TabIndex        =   2
      Top             =   3060
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   4320
      TabIndex        =   1
      Top             =   750
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   360
      Left            =   4320
      TabIndex        =   0
      Top             =   270
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDoctorTitle 
      Height          =   2985
      Left            =   90
      TabIndex        =   3
      Top             =   510
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   260
      RowHeightMax    =   260
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicDoctorTitleSet.frx":000C
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
      Caption         =   "    请设置各职称在挂号安排中医生姓名前显示的标识符。"
      Height          =   420
      Left            =   90
      TabIndex        =   4
      Top             =   120
      Width           =   4110
   End
End
Attribute VB_Name = "frmClinicDoctorTitleSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mblnChange As Boolean     '是否改变了

Public Function ShowMe(frmParent As Form) As Boolean
    '程序入口
    mblnOk = False
    On Error Resume Next
    Me.Show 1, frmParent
 
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Function TextIsValied(strText As String) As Boolean
    '输入文本是否有效
    Dim intCHeckLen As Integer
    
    intCHeckLen = 5
    If zlCommFun.StrIsValid(strText, intCHeckLen) = False Then Exit Function
    If InStr(strText, ",") > 0 Or InStr(strText, ";") Then
        MsgBox "标识符含有非法字符 ", vbInformation, gstrSysName
        Exit Function
    End If
    TextIsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    Dim i As Integer
    
    On Error GoTo ErrHandler
    '数据检查
    With vsfDoctorTitle
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("标识符"))) <> "" Then
                If TextIsValied(Trim(.TextMatrix(i, .ColIndex("标识符")))) = False Then
                    .Row = i: Exit Sub
                End If
            End If
        Next
    End With
    
    If SaveData() = False Then Exit Sub
    
    mblnOk = True
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
    '功能:保存编辑的内容
    '参数:
    '返回值:成功返回True,否则为False
    Dim i As Integer, str标识符 As String
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    '把所有设置的标识符做成一个串
    '格式：编码1,标识符1;编码2,标识符2;...
    With vsfDoctorTitle
        For i = .FixedRows To .Rows - 1
            str标识符 = str标识符 & ";"
            str标识符 = str标识符 & .TextMatrix(i, .ColIndex("编码")) & ","
            str标识符 = str标识符 & Trim(.TextMatrix(i, .ColIndex("标识符")))
        Next
    End With
    If str标识符 <> "" Then str标识符 = Mid(str标识符, 2)
    
    'Zl_专业技术职务_更新标识符
    strSQL = "Zl_专业技术职务_更新标识符("
    '    标识符_In In Varchar2 --格式：编码1,标识符1;编码2,标识符2;...
    strSQL = strSQL & "'" & str标识符 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(",';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    mblnChange = False
    
    On Error GoTo ErrHandler
    '23* 卫生技术人员（医疗）
    strSQL = "Select 编码, 名称, 标识符" & vbNewLine & _
            " From 专业技术职务" & vbNewLine & _
            " Where 编码 Like '23%' And 编码 <> '23'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.EOF Then
        MsgBox "    在“专业技术职务”表中没有找到编码以【23】开头表示“卫生技术人员（医疗）”的数据，请检查数据是否完整！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    With vsfDoctorTitle
        .Clear 1
        .Rows = rsTemp.RecordCount + 1
        .Editable = flexEDKbdMouse
        .GridLines = flexGridInset
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("编码")) = Nvl(rsTemp!编码)
            .TextMatrix(lngRow, .ColIndex("职称")) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("标识符")) = Nvl(rsTemp!标识符)
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        .Cell(flexcpBackColor, .FixedRows, .ColIndex("职称"), .Rows - 1, .ColIndex("职称")) = vbButtonFace
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True
End Sub

Private Sub vsfDoctorTitle_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsfDoctorTitle_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfDoctorTitle.ColIndex("职称") Then Cancel = True
End Sub

Private Sub vsfDoctorTitle_EnterCell()
    vsfDoctorTitle.EditCell
End Sub

Private Sub vsfDoctorTitle_GotFocus()
    With vsfDoctorTitle
        If .Row < .FixedRows And .Rows > .FixedRows Then .Row = .FixedRows
    End With
End Sub

Private Sub vsfDoctorTitle_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If TextIsValied(Trim(vsfDoctorTitle.EditText)) = False Then
        Cancel = True
    End If
End Sub
