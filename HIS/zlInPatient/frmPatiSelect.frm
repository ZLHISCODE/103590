VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPatiSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病人选择"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   10470
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vfgPati 
      Height          =   4155
      Left            =   2505
      TabIndex        =   7
      Top             =   480
      Width           =   7905
      _cx             =   13944
      _cy             =   7329
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
      SheetBorder     =   -2147483627
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
   Begin VB.CheckBox ChkSurety 
      Caption         =   "仅显示存在担保记录的病人"
      Height          =   180
      Left            =   2700
      TabIndex        =   5
      Top             =   4980
      Width           =   2610
   End
   Begin VB.CheckBox chkSect 
      Caption         =   "住院科室(不勾按病区)"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   2430
   End
   Begin VB.ComboBox cboSort 
      Height          =   300
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9245
      TabIndex        =   3
      Top             =   4875
      Width           =   1150
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7935
      TabIndex        =   2
      Top             =   4875
      Width           =   1150
   End
   Begin VB.ComboBox cboSect 
      Height          =   4140
      Left            =   45
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Text            =   "cboSect"
      Top             =   480
      Width           =   2400
   End
   Begin VB.Label lblSort 
      Caption         =   "缺省排序依据"
      Height          =   255
      Left            =   45
      TabIndex        =   6
      Top             =   4980
      Width           =   1215
   End
End
Attribute VB_Name = "frmPatiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mfrmParent As Form
Private mrsPati As New ADODB.Recordset
Private mint缺省排序 As Integer
Private mstrSort As String          '床号|住院号|病人ID|姓名|在院
Private mstrPatiTypeColor As String '病人颜色串  名称,颜色值|名称,颜色值----
Private mblnOk As Boolean

Public Function ShowMe(ByVal frmParent As Form) As Boolean
    Set mfrmParent = frmParent
    mblnOk = False
    Me.Show 1, mfrmParent
    ShowMe = mblnOk
End Function
Private Sub cboSect_Click()
    Dim strSQL As String, i As Integer, lngColor As Long, l As Integer
    Dim strSQL1 As String
    
    vfgPati.Clear
    If cboSect.ListIndex = -1 Then Exit Sub
    If mrsPati.State = adStateOpen Then mrsPati.Close
    Set mrsPati = New ADODB.Recordset
    On Error GoTo errHandle
    strSQL1 = ""
    strSQL1 = strSQL1 & IIf(chkSect.Value = 0, " And B.当前病区ID+0=[1]", " And B.出院科室ID+0 = [1]")
    strSQL1 = strSQL1 & IIf(ChkSurety.Value = 0, "", "  And EXISTS (SELECT 1 FROM 病人担保记录 WHERE 病人ID=B.病人ID AND 主页ID=B.主页ID AND ROWNUM<2) ")
    
    strSQL = "Select A.病人ID,B.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,A.家庭地址,B.出院病床 as 床位,Decode(B.出院日期,NULL,'√','') as 在院,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型" & _
            " From 病人信息 A,病案主页 B" & _
            " Where A.在院 = 1 And A.停用时间 is NULL And Nvl(B.主页ID,0)<>0" & _
            " And A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.出院日期 IS NULL " & strSQL1 & _
            " Order by " & Split(mstrSort, "|")(mint缺省排序) & " Desc"
    Set mrsPati = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(cboSect.ItemData(cboSect.ListIndex)))
    With vfgPati
        .Redraw = False: Set .DataSource = mrsPati
        
        If mrsPati.RecordCount > 0 Then
            .ColWidth(0) = 800
            .ColWidth(1) = 1000
            .ColWidth(2) = 800
            .ColWidth(3) = 500
            .ColWidth(4) = 2500
            .ColWidth(5) = 500
            .ColWidth(6) = 500
            .ColWidth(7) = 800
            DoEvents
            For i = 1 To .Rows - 1
                lngColor = GetPatiColor(.TextMatrix(i, 7))
                .Row = i
                For l = 0 To .Cols - 1
                    .Col = l
                    .CellForeColor = lngColor
                Next
            Next
        Else
            .Rows = 2
            .Cols = 2
        End If
        .RowHeight(-1) = 255
        .RowHeight(0) = 320
        .Row = 1: .TopRow = 1
        .Col = 0: .ColSel = .Cols - 1

        .Redraw = True
        If .Visible Then .SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboSect_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    If KeyAscii = 13 Then
        For i = 1 To cboSect.ListCount
            If cboSect.Text <> "" Then
                If cboSect.List(i) Like "*" & cboSect.Text & "*" Then
                    cboSect.ListIndex = i
                    Exit For
                End If
            End If
        Next
    End If
End Sub

Private Sub cboSort_Click()
    If cboSort.Visible And cboSort.ListIndex <> -1 Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "病人选择器排序", cboSort.ListIndex
        mint缺省排序 = cboSort.ListIndex
        Call cboSect_Click
    End If
End Sub

Private Sub chkSect_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    cboSect.Clear
    On Error GoTo errHandle
    
    If chkSect.Value = 0 Then '病区
        strSQL = "Select B.ID,B.编码,B.名称" & _
            " From (Select Distinct 病区ID From 床位状况记录 " & _
            " ) A,部门表 B Where A.病区ID=B.ID And (B.站点=[1] Or B.站点 is Null)" & _
            " Order by B.编码"
    Else '科室
        strSQL = "Select B.ID,B.编码,B.名称" & _
            " From (Select Distinct 科室ID From 床位状况记录 " & _
            " ) A,部门表 B Where A.科室ID=B.ID And (B.站点=[1] Or B.站点 is Null)" & _
            " Order by B.编码"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)

    With rsTmp
        Do While Not .EOF
            cboSect.AddItem !编码 & "-" & !名称
            cboSect.ItemData(cboSect.NewIndex) = !ID
            If !ID = UserInfo.部门ID Then cboSect.ListIndex = cboSect.NewIndex
            .MoveNext
        Loop
    End With
    If cboSect.ListCount > 0 And cboSect.ListIndex = -1 Then cboSect.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ChkSurety_Click()
    Call cboSect_Click
End Sub

Private Sub cmdCanc_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If vfgPati.Rows > 1 And vfgPati.TextMatrix(1, 0) <> "" Then
        mfrmParent.txtPatient.Text = "-" & vfgPati.TextMatrix(vfgPati.Row, 0)
        mblnOk = True
        Unload Me
    End If
End Sub

Private Sub vfgPati_DblClick()
    cmdOK_Click
End Sub

Private Sub vfgPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub vfgPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vfgPati.MouseRow = 0 Then
        vfgPati.MousePointer = 7
    Else
        vfgPati.MousePointer = 0
    End If
End Sub

Private Sub vfgPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    Dim lngColor As Long, i As Long, l As Long
    
    lngCol = vfgPati.MouseCol
    Debug.Print vfgPati.MouseCol
    If Button = 1 And vfgPati.MousePointer = 7 Then
        If vfgPati.TextMatrix(0, lngCol) = "" Then Exit Sub
        
        mrsPati.Sort = vfgPati.TextMatrix(0, lngCol) & IIf(vfgPati.ColData(lngCol) = 0, "", " DESC")
        vfgPati.Redraw = False
        Set vfgPati.DataSource = mrsPati
        If mrsPati.RecordCount > 0 Then
            For i = 1 To vfgPati.Rows - 1
                lngColor = GetPatiColor(vfgPati.TextMatrix(i, 7))
                DoEvents
                vfgPati.Row = i
                For l = 0 To vfgPati.Cols - 1
                    vfgPati.Col = l
                    vfgPati.CellForeColor = lngColor
                Next
            Next
            vfgPati.Row = 1: vfgPati.TopRow = 1
            vfgPati.Col = 0: vfgPati.ColSel = vfgPati.Cols - 1
        Else
            vfgPati.Rows = 2
            vfgPati.Cols = 2
        End If
        vfgPati.Redraw = True
        vfgPati.ColData(lngCol) = (vfgPati.ColData(lngCol) + 1) Mod 2
    End If
End Sub

Private Sub Form_Activate()
    vfgPati.SetFocus
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Call InitPatiType
    
    mstrSort = "床位|住院号|病人ID|姓名|在院|病人类型"
    For i = 0 To UBound(Split(mstrSort, "|"))
        cboSort.AddItem Split(mstrSort, "|")(i)
    Next
    mint缺省排序 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "病人选择器排序", 0))
    mint缺省排序 = IIf(mint缺省排序 < cboSort.ListCount, mint缺省排序, 0)
    cboSort.ListIndex = mint缺省排序
    If chkSect.Value = 1 Then
        Call chkSect_Click
    Else
        chkSect.Value = 1
    End If
End Sub

Private Sub lblSect_Click()
    cboSect.SetFocus
End Sub

Private Sub vfgPati_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyLeft Then
        If cboSect.ListIndex <> -1 Then
            If cboSect.ListIndex - 1 >= 0 Then
                cboSect.ListIndex = cboSect.ListIndex - 1
                vfgPati.Row = 1: vfgPati.Col = 0: vfgPati.ColSel = vfgPati.Cols - 1: vfgPati.SetFocus
            End If
        End If
    ElseIf KeyCode = vbKeyRight Then
        If cboSect.ListIndex <> -1 Then
            If cboSect.ListIndex + 1 <= cboSect.ListCount - 1 Then
                cboSect.ListIndex = cboSect.ListIndex + 1
                vfgPati.Row = 1: vfgPati.Col = 0: vfgPati.ColSel = vfgPati.Cols - 1: vfgPati.SetFocus
            End If
        End If
    End If
End Sub

Private Function InitPatiType() As Boolean
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    mstrPatiTypeColor = ""
    gstrSQL = "select 名称,颜色 from 病人类型"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取病人类型")
    Do Until rsTemp.EOF
        mstrPatiTypeColor = mstrPatiTypeColor & rsTemp!名称 & "," & Nvl(rsTemp!颜色, 0) & "|"
        rsTemp.MoveNext
    Loop
    If Len(mstrPatiTypeColor) > 0 Then
        mstrPatiTypeColor = Mid(mstrPatiTypeColor, 1, Len(mstrPatiTypeColor) - 1)
    Else
        mstrPatiTypeColor = "普通病人,0|医保病人,255"
    End If
    InitPatiType = True
    Exit Function
errH:
    mstrPatiTypeColor = "普通病人,0|医保病人,255"
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetPatiColor(ByVal strPatiType) As Long
    Dim arrType As Variant, i As Integer
    arrType = Split(mstrPatiTypeColor, "|")
    For i = LBound(arrType) To UBound(arrType)
        If Split(arrType(i), ",")(0) = strPatiType Then
            GetPatiColor = Split(arrType(i), ",")(1)
            Exit Function
        End If
    Next
End Function

