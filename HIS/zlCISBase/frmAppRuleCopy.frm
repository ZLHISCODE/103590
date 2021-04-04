VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmAppRuleCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "仪器规则复制"
   ClientHeight    =   7515
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7455
   Icon            =   "frmAppRuleCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin XtremeReportControl.ReportControl rptRule 
      Height          =   2865
      Left            =   150
      TabIndex        =   2
      Top             =   4485
      Width           =   7095
      _Version        =   589884
      _ExtentX        =   12515
      _ExtentY        =   5054
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnSort =   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "退出(&E)"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   3825
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "复制(&C)"
      Height          =   375
      Left            =   4710
      TabIndex        =   0
      Top             =   3825
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   3615
      Left            =   150
      TabIndex        =   3
      Top             =   105
      Width           =   7095
      _cx             =   12515
      _cy             =   6376
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   3795
      Top             =   4155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleCopy.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleCopy.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleCopy.frx":0B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleCopy.frx":10DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "注意：复制时将删除以前的已经设置了的规则。"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   195
      TabIndex        =   5
      Top             =   3900
      Width           =   3780
   End
   Begin VB.Label lbl源 
      Caption         =   "待复制规则"
      Height          =   210
      Left            =   180
      TabIndex        =   4
      Top             =   4245
      Width           =   2895
   End
End
Attribute VB_Name = "frmAppRuleCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mlng仪器ID As Long, mlng项目id As Long
Private rptCol As ReportColumn, rptRcd As ReportRecord, rptRow As ReportRow
Private Enum mColR
    性质 = 0: ID: 判断: 规则: 批范围: 多水平: 符合处理: 不符处理: 是否使用
End Enum

Public Sub ShowMe(ByVal lng仪器ID As Long, ByVal lng项目id As Long, frmMain As Form)

    If lng仪器ID = 0 Or lng项目id = 0 Then Exit Sub
    mlng仪器ID = lng仪器ID
    mlng项目id = lng项目id
    
    Me.Show vbModal, frmMain
    
End Sub

Private Function zlRefRule() As Long
    '功能：刷新装入当前仪器的规则
    Dim rsTemp As New ADODB.Recordset
    Dim blnCopy As Boolean
    gstrSql = "Select R.ID, R.上级id, R.性质, Decode(R.性质, 'Y', '符合: ', 'N', '不符: ', '') || R.判断 As 判断, B.名称 As 规则," & vbNewLine & _
            "       Decode(R.批范围, 1, '当前批', '近' || R.批范围 || '批') As 批范围, Decode(R.多水平, 1, '多', '') As 多水平," & vbNewLine & _
            "       Decode(Y结束, 0, '下一步', '结束') As 符合处理, Decode(N结束, 0, '下一步', '结束') As 不符处理,Decode(是否使用,1,'√','') as 是否使用" & vbNewLine & _
            "From (Select Level As 层次, ID, Nvl(上级id, 0) As 上级id, 性质, 判断, 规则id, 批范围, 多水平, Y结束, N结束, 是否使用" & vbNewLine & _
            "       From 检验仪器规则" & vbNewLine & _
            "       Where 仪器id = [1] And nvl(项目ID,0)=[2] " & vbNewLine & _
            "       Start With 仪器id = [1] And 上级id Is Null" & vbNewLine & _
            "       Connect By Prior ID = 上级id) R, 检验质控规则 B" & vbNewLine & _
            "Where R.规则id = B.ID" & vbNewLine & _
            "Order By R.层次, Decode(R.性质, '0', 0, '1', 1, 'Y', 2, 'N', 3, 1)"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlng仪器ID, mlng项目id)
    Err = 0: On Error GoTo 0
    Me.rptRule.Records.DeleteAll
    Me.rptRule.Populate
    With rsTemp
    
        Do While Not .EOF
            blnCopy = True
            If Val("" & !上级ID) = 0 Then
                Set rptRcd = Me.rptRule.Records.Add()
            Else
                Me.rptRule.Populate
                For Each rptRow In Me.rptRule.Rows
                    If Val(rptRow.Record(mColR.ID).Value) = Val("" & !上级ID) Then
                        Set rptRcd = rptRow.Record.Childs.Add()
                    End If
                Next
            End If
            If "" & !性质 = "1" Then
                rptRcd.AddItem("1").Icon = 3
            Else
                rptRcd.AddItem(CStr("" & !性质)).Icon = 2

            End If
            rptRcd.AddItem CStr(!ID)
            rptRcd.AddItem CStr(!判断)
            rptRcd.AddItem CStr("" & !规则)
            rptRcd.AddItem CStr("" & !批范围)
            rptRcd.AddItem CStr("" & !多水平)
            rptRcd.AddItem CStr("" & !符合处理)
            rptRcd.AddItem CStr("" & !不符处理)
            rptRcd.AddItem (CStr("" & !是否使用))
            rptRcd.Expanded = True
            .MoveNext
        Loop
    End With
    Me.rptRule.Populate
    OKButton.Enabled = blnCopy
    zlRefRule = Me.rptRule.Records.Count
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefRule = Me.rptRule.Records.Count
End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
  
    
    '-----------------------------------------------------
    '规则列表初始化
    With Me.rptRule
        .AutoColumnSizing = True
        Set rptCol = .Columns.Add(mColR.性质, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColR.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mColR.判断, "判断描述", 150, True): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.TreeColumn = True
        Set rptCol = .Columns.Add(mColR.规则, "判断规则", 82, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.批范围, "批范围", 45, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.多水平, "多水平", 45, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColR.符合处理, "符合处理", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.不符处理, "不符处理", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.是否使用, "是否使用", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        .SetImageList Me.ImgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列项目..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    Call zlRefRule
    
    '-----------------------------------------------------
    '初始化可复制的项目
    
    strSQL = "Select Distinct C.项目ID,Null as 选择,I.编码, I.名称 As 中文名, L.缩写 As 英文名" & vbNewLine & _
        "From 检验仪器项目 C, 检验项目 L, 检验报告项目 R, 诊疗项目目录 I, 检验质控品项目 A" & vbNewLine & _
        "Where A.项目id = C.项目id And C.项目id = L.诊治项目id And L.诊治项目id = R.报告项目id And R.诊疗项目id = I.ID And" & vbNewLine & _
        "      I.组合项目 <> 1 And L.项目类别 <> 2 And C.仪器id = [1] And C.项目ID<>[2] " & vbNewLine & _
        "Order By I.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng仪器ID, mlng项目id)
    Set Me.vfgList.DataSource = rsTmp
    Me.vfgList.ColWidth(0) = 0
    Me.vfgList.ColHidden(0) = True
    Me.vfgList.Cell(flexcpChecked, 1, 1, Me.vfgList.Rows - 1, 1) = flexUnchecked
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub OKButton_Click()
    Dim intRow As Integer, lngObj项目id As Long
    Dim strSQL As String
    Dim blnCheck As Boolean
    On Error GoTo errHandle
    
    With Me.vfgList
        For intRow = 1 To .Rows - 1
            If .Cell(flexcpChecked, intRow, 1) = flexChecked Then
                lngObj项目id = Val("" & .TextMatrix(intRow, 0))
                If lngObj项目id <> 0 Then
                    strSQL = "Zl_检验仪器规则_Copy(" & mlng仪器ID & "," & mlng项目id & "," & lngObj项目id & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                End If
                blnCheck = True
            End If
        Next
    End With
    
    If Not blnCheck Then
        MsgBox "请至少选择一个项目，然后再点复制！", vbInformation, Me.CancelButton
    Else
        Unload Me
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_Click()
    With Me.vfgList
        Debug.Print .MouseCol
        If .MouseCol = 1 Then
            .Cell(flexcpChecked, .Row, 1) = IIf(.Cell(flexcpChecked, .Row, 1) = flexUnchecked, flexChecked, flexUnchecked)
        End If
    End With
End Sub
