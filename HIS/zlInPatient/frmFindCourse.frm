VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFindCourse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找病人"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmFindCourse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3315
      TabIndex        =   3
      Top             =   660
      Width           =   1150
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3315
      TabIndex        =   2
      Top             =   135
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      Caption         =   "查找条件"
      Height          =   930
      Left            =   90
      TabIndex        =   4
      Top             =   60
      Width           =   3150
      Begin VB.TextBox txtSeekValue 
         Height          =   270
         Left            =   1290
         TabIndex        =   1
         Top             =   390
         Width           =   1725
      End
      Begin VB.ComboBox cboSeekKey 
         Height          =   300
         ItemData        =   "frmFindCourse.frx":000C
         Left            =   135
         List            =   "frmFindCourse.frx":0025
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   375
         Width           =   1020
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsgUpGrid 
      DragIcon        =   "frmFindCourse.frx":006F
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   4335
      _cx             =   7646
      _cy             =   2990
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
      TabBehavior     =   1
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
Attribute VB_Name = "frmFindCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOK As Boolean
Public mstrSeekKey As String
Public mstrSeekValue As String
Public mlng病区id As Long

Private Sub cboSeekKey_KeyPress(KeyAscii As Integer)
    cbo.AppendText cboSeekKey, KeyAscii
End Sub

Private Sub cmdCancel_Click()
    mstrSeekKey = ""
    mstrSeekValue = ""
    mlng病区id = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtSeekValue.Text) = "" Then
        MsgBox "请输入要查找的" & zlCommFun.GetNeedName(cboSeekKey.Text) & "！", vbInformation, gstrSysName
        txtSeekValue.SetFocus
        Exit Sub
    End If
    
    mstrSeekKey = zlCommFun.GetNeedName(cboSeekKey.Text)
    mstrSeekValue = Trim(txtSeekValue.Text)
    
    If mstrSeekKey <> "床号" And mlng病区id = 0 Then
        If Me.Height = 1500 Then Call LoadVfgData(vsgUpGrid, 2)
        If vsgUpGrid.Rows > 1 Then
            If vsgUpGrid.Rows > 2 Then
                If Me.Height = 3315 Then
                    mlng病区id = Val(vsgUpGrid.TextMatrix(vsgUpGrid.Row, vsgUpGrid.ColIndex("当前病区id")))
                Else
                    Me.Height = 3315
                    Exit Sub
                End If
            Else
                Me.Height = 1500
                mlng病区id = Val(vsgUpGrid.TextMatrix(vsgUpGrid.Row, vsgUpGrid.ColIndex("当前病区id")))
            End If
        Else
            MsgBox "你要查找的 " & mstrSeekKey & "=" & mstrSeekValue & " 的病人,不存在！", vbInformation, gstrSysName
            txtSeekValue.SetFocus
            Exit Sub
        End If
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    mblnOK = False
    If mstrSeekKey <> "" Then
        cbo.SetText cboSeekKey, mstrSeekKey
    Else
        cboSeekKey.ListIndex = 0
    End If
    
    If mstrSeekValue <> "" Then
        txtSeekValue.Text = mstrSeekValue
        zlControl.TxtSelAll txtSeekValue
    End If
    txtSeekValue.SetFocus
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        zlCommFun.PressKey vbKeyTab
'    End If
'End Sub

Private Sub Form_Load()
    Me.Height = 1500
    mlng病区id = 0
    Call LoadVfgData(vsgUpGrid, 1)
End Sub

Private Sub txtSeekValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mstrSeekKey = zlCommFun.GetNeedName(cboSeekKey.Text)
        mstrSeekValue = Trim(txtSeekValue.Text)
        If mstrSeekKey <> "床号" Then
            Call LoadVfgData(vsgUpGrid, 2)
            If vsgUpGrid.Rows > 1 Then
                If vsgUpGrid.Rows > 2 Then
                    Me.Height = 3315
                Else
                    Me.Height = 1500
                    mlng病区id = Val(vsgUpGrid.TextMatrix(vsgUpGrid.Row, vsgUpGrid.ColIndex("当前病区id")))
                End If
                zlCommFun.PressKey vbKeyTab
            Else
                MsgBox "你要查找的 " & mstrSeekKey & "=" & mstrSeekValue & " 的病人,不存在！", vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Sub txtSeekValue_KeyPress(KeyAscii As Integer)
    If zlCommFun.GetNeedName(cboSeekKey.Text) = "住院号" Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    '问题30342 by lesfeng 2010-06-01
    ElseIf zlCommFun.GetNeedName(cboSeekKey.Text) = "就诊卡号" Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If InStr("[]:：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub initvfgHeadTitle(ByVal vsGrid As VSFlexGrid)
    Dim strHead As String
    strHead = "姓名,500,4,1;住院号,800,1,1;次数,800,1,1;入院科室,1000,1,1;入院日期,1200,4,0;出院科室,1000,4,0;出院日期,1200,4,0;当前病区id,0,1,-1"
    Call SetVsFlexGridChangeHead(strHead, vsGrid, 0)
End Sub

Private Sub SetInitVfgFormat(ByVal vsGrid As VSFlexGrid)
    With vsGrid
        .ExplorerBar = flexExSortShowAndMove
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
    End With
End Sub

Private Sub LoadVfgData(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    Dim strSQL As String
    Dim strBillHead As String
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim strWhere As String
    Dim strValue As String
    Dim lngOutTime As Long
        
    Dim i As Long
    
    On Error GoTo errHandle
    
    If intFlag = 2 Then
        If mstrSeekKey = "住院号" Then
            strWhere = " And B." & mstrSeekKey & " = [1]"
        Else
            strWhere = " And A." & mstrSeekKey & " = [1]"
        End If
        
        strValue = mstrSeekValue
        
        lngOutTime = Val(zlDatabase.GetPara("出院天数", glngSys, 1132, "30"))
        strTemp = " And (B.出院日期 is null or B.出院日期>=" & IIf(lngOutTime <> 0, "Sysdate-[2]", "trunc(Sysdate)") & ")"
        
        strSQL = " Select A.姓名,B.住院号,B.主页id As 次数,B.入院科室id,B.入院日期,B.当前病区id,B.出院科室id,B.出院日期,C.名称 As 入院科室,D.名称 As 出院科室 " & _
                 "  From  病人信息 A,病案主页 B,部门表 C,部门表 D " & _
                 " Where A.病人id = B.病人id And B.入院科室id = C.ID And B.出院科室id = D.ID " & strWhere & strTemp
        Select Case mstrSeekKey
        Case "住院号"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strValue), lngOutTime)
        Case "医保号"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue, lngOutTime)
        Case "姓名"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue, lngOutTime)
        Case "就诊卡号"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue, lngOutTime)
        Case "身份证号"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue, lngOutTime)
        Case "IC卡号"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue, lngOutTime)
        End Select
    End If
            
    With vsGrid
        .Clear
        Call initvfgHeadTitle(vsGrid)
        If intFlag = 2 Then
            .Rows = IIf(rsTemp.EOF, 0, rsTemp.RecordCount) + 1
            If Not rsTemp.EOF Then
                For i = 1 To .Rows - 1
                    .TextMatrix(i, .ColIndex("姓名")) = IIf(IsNull(rsTemp!姓名), "", rsTemp!姓名)
                    .TextMatrix(i, .ColIndex("住院号")) = IIf(IsNull(rsTemp!住院号), 0, rsTemp!住院号)
                    .TextMatrix(i, .ColIndex("次数")) = IIf(IsNull(rsTemp!次数), 0, rsTemp!次数)
                    .TextMatrix(i, .ColIndex("入院科室")) = IIf(IsNull(rsTemp!入院科室), "", rsTemp!入院科室)
                    .TextMatrix(i, .ColIndex("入院日期")) = IIf(IsNull(rsTemp!入院日期), "", rsTemp!入院日期)
                    .TextMatrix(i, .ColIndex("出院科室")) = IIf(IsNull(rsTemp!出院科室), "", rsTemp!出院科室)
                    .TextMatrix(i, .ColIndex("出院日期")) = IIf(IsNull(rsTemp!出院日期), "", rsTemp!出院日期)
                    .TextMatrix(i, .ColIndex("当前病区id")) = IIf(IsNull(rsTemp!当前病区ID), 0, rsTemp!当前病区ID)
                    rsTemp.MoveNext
                Next
            End If
        End If
    End With
    Call SetInitVfgFormat(vsGrid)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsgUpGrid_DblClick()
    If vsgUpGrid.Rows > 1 Then
        mlng病区id = Val(vsgUpGrid.TextMatrix(vsgUpGrid.Row, vsgUpGrid.ColIndex("当前病区id")))
        Call cmdOK_Click
    End If
End Sub


