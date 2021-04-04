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
   Begin VSFlex8Ctl.VSFlexGrid vsfPatient 
      Height          =   4185
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   7875
      _cx             =   13891
      _cy             =   7382
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin VB.ComboBox cbo缺省排序 
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
      TabIndex        =   2
      Top             =   4875
      Width           =   1150
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7665
      TabIndex        =   1
      Top             =   4875
      Width           =   1150
   End
   Begin VB.ComboBox cboSect 
      Height          =   4140
      Left            =   45
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "cboSect"
      Top             =   480
      Width           =   2400
   End
   Begin VB.Label lbl缺省排序 
      Caption         =   "缺省排序依据"
      Height          =   255
      Left            =   45
      TabIndex        =   5
      Top             =   4980
      Width           =   1215
   End
   Begin VB.Label lblSect 
      AutoSize        =   -1  'True
      Caption         =   "住院科室"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   210
      Width           =   720
   End
End
Attribute VB_Name = "frmPatiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mfrmParent As Form
Private mrsPati As New ADODB.Recordset
Private mint缺省排序 As Integer
Private mstrSort As String          '床号|住院号|病人ID|姓名|在院
Private mblnSort As Boolean
'81407,调整选择器界面控件字体大小
Public mbytSize As Byte '字体：0-小字体,1-大字体;小字体为9号字,大字体为12号字

Private Sub cboSect_Click()
    Dim strSQL As String, i As Integer, lngColor As Long, l As Integer
    
    vsfPatient.Clear 1
    If cboSect.ListIndex = -1 Then Exit Sub
    If mrsPati.State = adStateOpen Then mrsPati.Close
    
    On Error GoTo errHandle
'    If Not gblnAllowOut Then
'        '当前在院病人
'        strSQL = " Select A.病人id, A.住院号, A.姓名,A.性别,A.家庭地址, A.当前床号 As 床位,'√' As 在院,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型" & _
'                 " From 病人信息 A, 病案主页 B" & _
'                 " Where A.在院 = 1 And A.病人id = B.病人id And A.主页ID = B.主页id And A.停用时间 Is Null And B.出院日期 Is Null And " & _
'                 " B.出院科室id+0 =[1]" & _
'                 " Order by " & Split(mstrSort, "|")(mint缺省排序) & " Desc"
'    Else
        '住(过)院病人
        '58842,刘鹏飞,2013-02-25,在院病人读取(从在院病人中读取)
        strSQL = "Select A.病人ID,A.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,A.家庭地址,B.出院病床 as 床位,Decode(B.出院日期,NULL,'√','') as 在院,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型" & _
                " From 病人信息 A,病案主页 B" & _
                " Where A.停用时间 is NULL And Nvl(B.主页ID,0)<>0" & _
                " And A.病人ID=B.病人ID And A.主页ID=B.主页ID And (Nvl(A.在院,0) = 1 Or Exists (Select 1 From 病案主页 Where 病人ID=A.病人ID And Nvl(主页ID,0)=0 And Nvl(病人性质,0)=0)) " & _
                " And A.当前科室ID=[1]" & _
                " Order by " & Split(mstrSort, "|")(mint缺省排序) & " Desc"

    'End If
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(cboSect.ItemData(cboSect.ListIndex)))
    With vsfPatient
        '86344:李南春,2015/7/8,msh换成vsfFlexGrid
        .Rows = 2
        .Cols = 8
        .TextMatrix(0, 0) = "病人ID"
        .TextMatrix(0, 1) = "住院号"
        .TextMatrix(0, 2) = "姓名"
        .TextMatrix(0, 3) = "性别"
        .TextMatrix(0, 4) = "家庭地址"
        .TextMatrix(0, 5) = "床位"
        .TextMatrix(0, 6) = "在院"
        .TextMatrix(0, 7) = "病人类型"
        For i = 0 To 7
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        Do While Not mrsPati.EOF
            .TextMatrix(.Rows - 1, 0) = NVL(mrsPati!病人ID)
            .TextMatrix(.Rows - 1, 1) = NVL(mrsPati!住院号)
            .TextMatrix(.Rows - 1, 2) = NVL(mrsPati!姓名)
            .TextMatrix(.Rows - 1, 3) = NVL(mrsPati!性别)
            .TextMatrix(.Rows - 1, 4) = NVL(mrsPati!家庭地址)
            .TextMatrix(.Rows - 1, 5) = NVL(mrsPati!床位)
            .TextMatrix(.Rows - 1, 6) = NVL(mrsPati!在院)
            .TextMatrix(.Rows - 1, 7) = NVL(mrsPati!病人类型)
            .Rows = .Rows + 1
            mrsPati.MoveNext
        Loop
        
'        .Redraw = flexRDBuffered
        
        If mrsPati.RecordCount > 0 Then
            '自动调整MSHFlexGrid表格的各列宽度
            Call zlControl.MshSetColWidth(vsfPatient, Me)
            For i = 1 To .Rows - 1
                lngColor = GetPatiColor(.TextMatrix(i, 7))
'                .Row = i
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = lngColor
'                For l = 0 To .Cols - 1
'                    .Col = l
'                    .CellForeColor = lngColor
'                Next
            Next
            .Rows = .Rows - 1
        Else
            .Rows = 2
            .Cols = 2
        End If

'        .RowHeight(0) = 320
'        .Row = 1: .TopRow = 1
'        .Redraw = flexRDDirect
        If .Visible And .Enabled = True Then .SetFocus
        If .Cols > 2 Then
        Select Case mint缺省排序
            Case 0
                .Col = 5
            Case 1
                .Col = 1
            Case 2
                .Col = 0
            Case 3
                .Col = 2
            Case 4
                .Col = 6
            Case 5
                .Col = 7
        End Select
        .Sort = flexSortGenericDescending
        End If
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetColor()
    Dim i As Integer, lngColor As Long
    With vsfPatient
        For i = 1 To .Rows - 1
            lngColor = GetPatiColor(.TextMatrix(i, 7))
'                .Row = i
            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = lngColor
'                For l = 0 To .Cols - 1
'                    .Col = l
'                    .CellForeColor = lngColor
'                Next
        Next
    End With
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

Private Sub cbo缺省排序_Click()
    If cbo缺省排序.Visible And cbo缺省排序.ListIndex <> -1 Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "病人选择器排序", cbo缺省排序.ListIndex
        mint缺省排序 = cbo缺省排序.ListIndex
        Call cboSect_Click
    End If
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If vsfPatient.Rows > 1 Then
        If vsfPatient.TextMatrix(1, 0) <> "" Then
            mfrmParent.txtPatient.Text = "-" & vsfPatient.TextMatrix(vsfPatient.Row, 0)
            Unload Me
        End If
    End If
End Sub

Private Sub vsfPatient_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsfPatient_DblClick()
    If vsfPatient.MouseRow > 0 Then cmdOK_Click
End Sub

Private Sub vsfPatient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub vsfPatient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsfPatient.MouseRow = 0 Then
        vsfPatient.MousePointer = 99
    Else
        vsfPatient.MousePointer = 0
    End If
End Sub

Private Sub vsfPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long

    lngCol = vsfPatient.MouseCol

    If Button = 1 And vsfPatient.MousePointer = 99 Then
        If vsfPatient.TextMatrix(0, lngCol) = "" Then Exit Sub
        vsfPatient.Col = lngCol
        If mblnSort Then
            vsfPatient.Sort = flexSortGenericAscending
            mblnSort = False
        Else
            vsfPatient.Sort = flexSortGenericDescending
            mblnSort = True
        End If
'        mrsPati.Sort = vsfPatient.TextMatrix(0, lngCol) & IIf(vsfPatient.ColData(lngCol) = 0, "", " DESC")
'        Set vsfPatient.DataSource = mrsPati

        vsfPatient.ColData(lngCol) = (vsfPatient.ColData(lngCol) + 1) Mod 2
    End If
End Sub

Private Sub Form_Activate()
    vsfPatient.SetFocus
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    
    '设置界面控件字体大小及位置
    Call SetFontSize(Me, mbytSize)
    If mbytSize = 1 Then
        cboSect.Height = cboSect.Height + 240
        vsfPatient.Height = vsfPatient.Height + 120
        lbl缺省排序.Width = lbl缺省排序.Width + 320
        cmdOK.Move cmdOK.Left - 2 * (1500 - cmdOK.Width), cmdOK.Top - 50, 1500, 420
        cmdCanc.Move cmdCanc.Left - (1800 - cmdCanc.Width), cmdCanc.Top - 50, 1500, 420
    End If
    cbo缺省排序.Left = lbl缺省排序.Left + lbl缺省排序.Width + 50
    
    Call InitPatiType
    
    mstrSort = "床位|住院号|病人ID|姓名|在院|病人类型"
    For i = 0 To UBound(Split(mstrSort, "|"))
        cbo缺省排序.AddItem Split(mstrSort, "|")(i)
    Next
    mint缺省排序 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "病人选择器排序", 0))
    mint缺省排序 = IIf(mint缺省排序 < cbo缺省排序.ListCount, mint缺省排序, 0)
    cbo缺省排序.ListIndex = mint缺省排序
    
    cboSect.Clear
    
    On Error GoTo errHandle
    'by lesfeng 2010-03-08 性能优化
    strSQL = "Select B.ID,B.编码,B.名称" & _
        " From (Select Distinct 科室ID From 床位状况记录 " & _
        " ) A,部门表 B Where A.科室ID=B.ID And (B.站点=[1] Or B.站点 is Null)" & _
        " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)

    With rsTmp
        Do While Not .EOF
            cboSect.AddItem !编码 & "-" & !名称
            cboSect.ItemData(cboSect.NewIndex) = !ID
            If !ID = UserInfo.部门ID Then cboSect.ListIndex = cboSect.NewIndex
            .MoveNext
        Loop
    End With
    vsfPatient.AllowUserResizing = flexResizeColumns
    If cboSect.ListCount > 0 And cboSect.ListIndex = -1 Then cboSect.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lblSect_Click()
    cboSect.SetFocus
End Sub

Private Sub vsfPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyLeft Then
        If cboSect.ListIndex <> -1 Then
            If cboSect.ListIndex - 1 >= 0 Then
                cboSect.ListIndex = cboSect.ListIndex - 1
                vsfPatient.Row = 1: vsfPatient.Col = 0: vsfPatient.SetFocus
            End If
        End If
    ElseIf KeyCode = vbKeyRight Then
        If cboSect.ListIndex <> -1 Then
            If cboSect.ListIndex + 1 <= cboSect.ListCount - 1 Then
                cboSect.ListIndex = cboSect.ListIndex + 1
                vsfPatient.Row = 1: vsfPatient.Col = 0: vsfPatient.SetFocus
            End If
        End If
    End If
End Sub

Private Sub SetFontSize(ByVal objForm As Object, ByVal bytSize As Byte)
    '设置界面控件字体大小
    '入参:
    '   objForm-窗体对象
    '   bytSize-字体大小: 0-小字体,1-大字体;小字体为9号字,大字体为12号字
    Dim objCtl As Control
    
    On Error Resume Next
    objForm.Font.Size = IIf(bytSize = 1, 12, 9)
    For Each objCtl In objForm.Controls
        '0-小字体,1-大字体;小字体为9号字,大字体为12号字
        objCtl.Font.Size = IIf(bytSize = 1, 12, 9)
    Next
End Sub
