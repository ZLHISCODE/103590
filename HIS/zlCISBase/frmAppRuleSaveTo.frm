VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAppRuleSaveTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "另存为范例"
   ClientHeight    =   3585
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6195
   Icon            =   "frmAppRuleSaveTo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4830
      TabIndex        =   3
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CommandButton cmdSaveTo 
      Caption         =   "保存(&S)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4830
      TabIndex        =   2
      Top             =   2715
      Width           =   1245
   End
   Begin VB.TextBox txt范例名 
      Height          =   300
      Left            =   960
      MaxLength       =   30
      TabIndex        =   1
      Top             =   2745
      Width           =   3045
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2505
      Left            =   105
      TabIndex        =   4
      Top             =   105
      Width           =   5970
      _cx             =   10530
      _cy             =   4419
      Appearance      =   2
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
      Cols            =   3
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
   Begin VB.Label lbl水平数 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "适用于每分析批检测水平数为**的仪器"
      Height          =   180
      Left            =   960
      TabIndex        =   5
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label lbl范例名 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "范例名(&N)"
      Height          =   180
      Left            =   105
      TabIndex        =   0
      Top             =   2805
      Width           =   810
   End
End
Attribute VB_Name = "frmAppRuleSaveTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum mCol
    范例名 = 0: 水平数: 适用于
End Enum

Private mlngDevId As Long           '仪器id
Private mblnOK As Boolean
Private mlngGroupID As Long         '分组ID
'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long

'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Public Function ShowMe(frmParent As Form, lngDevId As Long, lngGroupID As Long) As Boolean
    '功能：刷新装入指定仪器
    Dim rsTemp As New ADODB.Recordset
    
    mlngDevId = lngDevId
    mlngGroupID = lngGroupID
    
    gstrSql = "Select Decode(A.质控水平数, Null, 1, 0, 1, A.质控水平数) As 水平数 From 检验仪器 A Where A.ID = [1]"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevId)
    
    Me.lbl水平数.Caption = "适用于每分析批检测水平数为" & rsTemp!水平数 & "的仪器"
    Me.txt范例名.Text = "新范例" & lngDevId & "(N=" & rsTemp!水平数 & ")"
    
    gstrSql = "Select Distinct 范例名, 水平数, '适用于每分析批检测水平数为' || 水平数 || '的仪器...' As 适用于" & vbNewLine & _
        "From 检验质控范则" & vbNewLine & _
        "Order By 范例名"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Set Me.vfgList.DataSource = rsTemp
    Me.vfgList.ColWidth(mCol.水平数) = 0
    Me.vfgList.ColHidden(mCol.水平数) = True
    
    mblnOK = False
    Me.Show vbModal, frmParent
    ShowMe = mblnOK
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False
End Function

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdSaveTo_Click()
    If Trim(Me.txt范例名.Text) = "" Then
        MsgBox "请输入范例名！", vbInformation, gstrSysName
        Me.txt范例名.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt范例名.Text), vbFromUnicode)) > Me.txt范例名.MaxLength Then
        MsgBox "范例名超长（最多" & Me.txt范例名.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt范例名.SetFocus: Exit Sub
    End If
    
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(lngCount, mCol.范例名)) = Trim(Me.txt范例名.Text) Then
                If MsgBox("真的要替换范例“" & .TextMatrix(.Row, mCol.范例名) & "”吗", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        Next
    End With
    
    gstrSql = "Zl_检验质控范则_Edit(1,'" & Trim(Me.txt范例名.Text) & "'," & mlngDevId & "," & mlngGroupID & ")"
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    mblnOK = True
    Unload Me: Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt范例名_GotFocus()
    Me.txt范例名.SelStart = 0: Me.txt范例名.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt范例名_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfgList_DblClick()
    With Me.vfgList
        If .Row < .FixedRows Then Exit Sub
        Me.txt范例名.Text = .TextMatrix(.Row, mCol.范例名)
    End With
End Sub

Private Sub vfgList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then Exit Sub
    With Me.vfgList
        If .Row < .FixedRows Then Exit Sub
        If MsgBox("真的要删除范例“" & .TextMatrix(.Row, mCol.范例名) & "”吗", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSql = "Zl_检验质控范则_Edit(2,'" & Trim(.TextMatrix(.Row, mCol.范例名)) & "')"
    End With
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Me.vfgList.RemoveItem Me.vfgList.Row
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
