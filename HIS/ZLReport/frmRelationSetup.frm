VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelationSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关联报表参数对照 - 手术医师费用汇总"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   Icon            =   "frmRelationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4245
      TabIndex        =   5
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPars 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5295
      _cx             =   9340
      _cy             =   2566
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRelationSetup.frx":6852
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
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   5535
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5535
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   -45
         TabIndex        =   1
         Top             =   810
         Width           =   8430
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   390
         Picture         =   "frmRelationSetup.frx":690E
         Top             =   75
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "请设置您关联的报表的参数来源，关联后查询此报表时，双击此元素将根据设置的参数来源作为关联报表的参数来执行关联报表。"
         ForeColor       =   &H00800000&
         Height          =   600
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   4260
      End
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3000
      Left            =   1200
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   5292
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      PathSeparator   =   "."
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmRelationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngReportID As Long
Private mstrData As String  '当前表格元素对应的数据源名称
Private mfrmParent As Object
Private mobjReport As Report
Private mstrCaption As String
Private mobjRelations As RPTRelations
Private Enum enum_Col
    col参数名 = 0
    col参数值来源 = 1
End Enum
Private mblnOK As Boolean
Private mlngType As Long '0-汇总表项，1-任意表项，2-标签

Public Function ShowMe(frmParent As Object, ByVal lngReportID As Long, ByVal strData As String, _
    ByRef objReport As Report, ByVal strCaption As String, _
    ByRef objRelations As RPTRelations, ByVal lngType As Long) As Boolean
    
    mlngReportID = lngReportID
    mstrData = strData
    mstrCaption = strCaption
    mlngType = lngType
    Set mfrmParent = frmParent
    Set mobjReport = objReport
    Set mobjRelations = objRelations
    
    Me.Show 1, frmParent
    Set objReport = mobjReport
    Set objRelations = mobjRelations
    ShowMe = mblnOK
End Function

Private Sub LoadPars()
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    
    strSQL = "Select distinct a.名称,b.报表id " & vbNewLine & _
            "From zlRPTPars A, zlRPTDatas B" & vbNewLine & _
            "Where a.源id = b.Id And b.报表id = [1]"
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, mlngReportID)
    With vsPars
        .Redraw = False
        .Rows = 1
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, col参数名) = rsTmp!名称 & ""
            '已绑定的参数对照
            .TextMatrix(.Rows - 1, col参数值来源) = mobjRelations.Item(rsTmp!报表id & "_" & rsTmp!名称).参数值来源
            rsTmp.MoveNext
        Loop
        .Redraw = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    If tvw.Visible Then
        tvw.Visible = False
    Else
        Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    For i = mobjRelations.count To 1 Step -1
        If mobjRelations.Item(i).关联报表ID = mlngReportID Then
            mobjRelations.Remove i
        End If
    Next
    With vsPars
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col参数值来源) <> "" Then
                If GetItem("_" & .TextMatrix(i, col参数名)) = False Then
                    mobjRelations.Add mlngReportID, .TextMatrix(i, col参数名), .TextMatrix(i, col参数值来源), mstrCaption, 0
                End If
            End If
        Next
    End With
    mblnOK = True
    Unload Me
End Sub

Private Function GetItem(ByVal strKey As String) As Boolean
    Err.Clear
    On Error GoTo ErrHand
    Dim var As RPTRelation
    Set var = mobjRelations.Item(strKey)
    If var.Key = "" Then
        GetItem = False
    Else
        GetItem = True
    End If
    Exit Function
ErrHand:
    Err.Clear
End Function

Private Sub Form_Load()
    vsPars.Editable = flexEDKbdMouse
    Me.Caption = "关联报表参数对照 - " & mstrCaption
    Call LoadPars
    If mlngType = Val("2-标签") Then
        Call CopySubTree(mfrmParent.tvwSQL)
    Else
        Call CopySubTree(mfrmParent.tvw)
    End If
    mblnOK = False
End Sub

Private Sub Form_Resize()
    vsPars.Height = vsPars.Rows * vsPars.RowHeightMin + 100
    cmdOK.Top = vsPars.Height + vsPars.Top + 100
    cmdCancel.Top = cmdOK.Top
    Me.Height = cmdCancel.Top + cmdCancel.Height + 500
End Sub

Private Sub vsPars_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewCol = col参数名 Then
        vsPars.Editable = flexEDNone
    Else
        vsPars.Editable = flexEDKbdMouse
        vsPars.ColComboList(NewCol) = "..."
    End If
    tvw.Visible = False
End Sub

Private Sub CopySubTree(objtvw As Object)
    Dim objNode As Object, tmpNode As Object
    Dim objPar As RPTPar
    Dim objData As RPTData
    Dim strTmp As String
    
    For Each objNode In objtvw.Nodes
        If objNode.Text = mstrData And objNode.Children <> 0 And objNode.Key <> "Root" Then Exit For
    Next
    
    tvw.Nodes.Clear
    Set tvw.ImageList = objtvw.ImageList
    
    Set tmpNode = tvw.Nodes.Add(, , objNode.Key, objNode.Text, objNode.Image, objNode.SelectedImage)
    tmpNode.Expanded = True
    tmpNode.Selected = True
    
    Set objNode = objNode.Child
    Do While Not objNode Is Nothing
        If Not IsType(Val(objNode.Tag), adLongVarBinary) Then
            Set tmpNode = tvw.Nodes.Add(objNode.Parent.Key, 4, objNode.Key, objNode.Text, objNode.Image, objNode.SelectedImage)
            tmpNode.Tag = objNode.Tag
        End If
        Set objNode = objNode.Next
    Loop
    For Each objData In mobjReport.Datas
        For Each objPar In objData.Pars
            If InStr(strTmp & ",", "," & objPar.名称 & ",") = 0 Then
                '先加根节点
                If strTmp = "" Then
                    Set tmpNode = tvw.Nodes.Add(, , "ParsRoot", "参数列表", "ParsRoot", "ParsRoot")
                    tmpNode.Expanded = True
                End If
                Set tmpNode = tvw.Nodes.Add("ParsRoot", 4, "_" & objPar.名称, objPar.名称, "Pars", "Pars")
                strTmp = strTmp & "," & objPar.名称
            End If
        Next
    Next
End Sub

Private Sub vsPars_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    SetParent tvw.hwnd, 0
    tvw.Top = Me.Top + vsPars.Top + (Row + 1) * vsPars.RowHeightMin + 400
    tvw.Left = Me.Left + vsPars.Left + vsPars.ColWidth(0) + (Col - 1) * vsPars.ColWidth(1) + 80
    tvw.ZOrder
    tvw.Visible = Not tvw.Visible
    vsPars.SetFocus
End Sub


Private Sub tvw_LostFocus()
    tvw.Visible = False
    vsPars.SetFocus
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key <> "Root" And Node.Key <> "ParsRoot" And Node.Children = 0 Then
        If Node.Parent.Text = "参数列表" Then
            vsPars.TextMatrix(vsPars.Row, vsPars.Col) = "=" & LevelText(Node)
        Else
            vsPars.TextMatrix(vsPars.Row, vsPars.Col) = Node.Parent.Text & "." & LevelText(Node)
        End If
    Else
        vsPars.TextMatrix(vsPars.Row, vsPars.Col) = ""
    End If
    tvw.Visible = False
    vsPars.SetFocus
End Sub

Private Sub vsPars_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        If MsgBox("是否删除本参数的对照信息，删除后如果不进行对照，进行关联查询时将弹出弹出输入框，临时决定参数值。", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            vsPars.TextMatrix(vsPars.Row, col参数值来源) = ""
        End If
    End If
End Sub


