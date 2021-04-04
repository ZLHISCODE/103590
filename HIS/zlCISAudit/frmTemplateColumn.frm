VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmTemplateColumn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择列项"
   ClientHeight    =   5850
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4935
   Icon            =   "frmTemplateColumn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   3870
      Index           =   1
      Left            =   225
      ScaleHeight     =   3870
      ScaleWidth      =   3345
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   3345
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   750
         Left            =   255
         TabIndex        =   2
         Top             =   255
         Width           =   2550
         _cx             =   4498
         _cy             =   1323
         Appearance      =   0
         BorderStyle     =   0
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   165
      TabIndex        =   12
      Top             =   5130
      Width           =   4665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3720
      TabIndex        =   10
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2550
      TabIndex        =   9
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "隐藏(&H)"
      Height          =   350
      Left            =   3720
      TabIndex        =   8
      Top             =   1980
      Width           =   1100
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "显示(&S)"
      Height          =   350
      Left            =   3720
      TabIndex        =   7
      Top             =   1605
      Width           =   1100
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2850
      TabIndex        =   4
      Top             =   4815
      Width           =   675
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "下移(&D)"
      Height          =   350
      Left            =   3720
      TabIndex        =   6
      Top             =   1215
      Width           =   1100
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "上移(&U)"
      Height          =   350
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   1100
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "详细信息(&T):"
      Height          =   180
      Left            =   210
      TabIndex        =   0
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "所选栏的宽度(像素)(&W):"
      Height          =   180
      Left            =   210
      TabIndex        =   3
      Top             =   4875
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "选择您想显示的项目的详细信息。"
      Height          =   180
      Left            =   180
      TabIndex        =   11
      Top             =   195
      Width           =   2700
   End
End
Attribute VB_Name = "frmTemplateColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mclsVsf As clsVsf
Private mobjClsVsf As clsVsf
Private mintHideCols As Integer

Public Function ShowColumn(ByVal frmMain As Object, ByRef objClsVsf As clsVsf) As Boolean
    
    mblnOK = False
    
    Set mobjClsVsf = objClsVsf
    
    If ExecuteCommand("初始控件") = False Then Exit Function
    If ExecuteCommand("初始数据") = False Then Exit Function
    
    Me.Show 1, frmMain

    If mblnOK Then
        Set objClsVsf = mobjClsVsf
        ShowColumn = mblnOK
    End If
    Set mobjClsVsf = Nothing
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String

    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
        Case "初始控件"
        
            Set mclsVsf = New clsVsf
            
            With mclsVsf
            
                Call .Initialize(Me.Controls, vsf, True, True, GetImageList(16))
                Call .ClearColumn
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTBoolean, "", "[选择]", False)
                                    
                Call .AppendColumn("AllowColEdit", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("AllowColSort", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("AllowColDrag", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("SysHidden", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("DropColData", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("EditControl", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("AllowColResize", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                
                Call .AppendColumn("宽度", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("名称", 1080, flexAlignLeftCenter, flexDTString, "", "", True)
                
                Call .InitializeEdit(True, False, False)
                Call .InitializeEditColumn(0, True, vbVsfEditCheck)
    
            End With
            
            vsf.Move 15, 15, picPane(1).Width - 30, picPane(1).Height - 30
            
            vsf.RowHidden(0) = True
        '--------------------------------------------------------------------------------------------------------------
        Case "初始数据"
            
            With mobjClsVsf.VsfBody
                
                mintHideCols = 0
                For intRow = 0 To .Cols - 1
                    strTmp = .TextMatrix(0, intRow)
                    If mobjClsVsf.SysHidden(intRow) = False Then
                    
                        strTmp = .TextMatrix(0, intRow)
                        If strTmp = "" Then
                            strTmp = .Cell(flexcpData, 0, intRow, 0, intRow)
                            If InStr(strTmp, "[") > 0 Then strTmp = Mid(strTmp, 2, Len(strTmp) - 2)
                        End If
                    
                        If vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("名称")) <> "" Then vsf.Rows = vsf.Rows + 1
    
                        vsf.TextMatrix(vsf.Rows - 1, 0) = IIf(.ColHidden(intRow), 0, 1)
                        vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("宽度")) = .ColWidth(intRow) / 15
                        vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("名称")) = strTmp
                        vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("AllowColEdit")) = mobjClsVsf.AllowColEdit(intRow)
                        vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("AllowColSort")) = mobjClsVsf.AllowColSort(intRow)
                        vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("AllowColDrag")) = mobjClsVsf.AllowColDrag(intRow)
                        vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("SysHidden")) = mobjClsVsf.SysHidden(intRow)
                        vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("DropColData")) = mobjClsVsf.DropColData(intRow)
                        vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("EditControl")) = mobjClsVsf.EditControl(intRow)
                        vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("AllowColResize")) = mobjClsVsf.AllowColResize(intRow)
                        
                    Else
                        mintHideCols = mintHideCols + 1
                    End If
        
                Next
                
                vsf.Col = vsf.ColIndex("名称")
                Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
                
            End With
    End Select
    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    If vsf.Row < vsf.Rows - 1 Then
        If mclsVsf.MoveRow(vsf.Row, 1) Then
            vsf.Row = vsf.Row + 1
        End If
    End If
End Sub

Private Sub cmdHide_Click()
    vsf.TextMatrix(vsf.Row, 0) = 0
End Sub

Private Sub cmdOK_Click()
    Dim intRow As Integer
    Dim intIndex As Integer
    
    With mobjClsVsf.VsfBody
        '保存设置
        
        For intRow = 1 To vsf.Rows - 1
            
            intIndex = .ColIndex(vsf.TextMatrix(intRow, vsf.ColIndex("名称")))
            '应用宽度
            .ColWidth(intIndex) = Val(vsf.TextMatrix(intRow, vsf.ColIndex("宽度"))) * 15
            
            '应用显示/隐藏
            .ColHidden(intIndex) = IIf(Abs(Val(vsf.TextMatrix(intRow, 0))) = 1, False, True)
            
            '应用位置
            .ColPosition(intIndex) = intRow - 1 + mintHideCols

            mobjClsVsf.AllowColEdit(intRow - 1 + mintHideCols) = vsf.TextMatrix(intRow, vsf.ColIndex("AllowColEdit"))
            mobjClsVsf.SysHidden(intRow - 1 + mintHideCols) = vsf.TextMatrix(intRow, vsf.ColIndex("SysHidden"))
            mobjClsVsf.AllowColSort(intRow - 1 + mintHideCols) = vsf.TextMatrix(intRow, vsf.ColIndex("AllowColSort"))
            mobjClsVsf.AllowColDrag(intRow - 1 + mintHideCols) = vsf.TextMatrix(intRow, vsf.ColIndex("AllowColDrag"))
            mobjClsVsf.EditControl(intRow - 1 + mintHideCols) = vsf.TextMatrix(intRow, vsf.ColIndex("EditControl"))
            mobjClsVsf.DropColData(intRow - 1 + mintHideCols) = vsf.TextMatrix(intRow, vsf.ColIndex("DropColData"))
            mobjClsVsf.AllowColResize(intRow - 1 + mintHideCols) = vsf.TextMatrix(intRow, vsf.ColIndex("AllowColResize"))
            
        Next

    End With
    
    mblnOK = True
    
    Unload Me
End Sub

Private Sub cmdShow_Click()
    vsf.TextMatrix(vsf.Row, 0) = 1
End Sub

Private Sub cmdUp_Click()
    If vsf.Row > 1 Then
        If mclsVsf.MoveRow(vsf.Row, -1) Then
            vsf.Row = vsf.Row - 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf = Nothing
'    Set mobjClsVsf = Nothing
End Sub

Private Sub txt_Change()
    vsf.TextMatrix(vsf.Row, vsf.ColIndex("宽度")) = Val(txt.Text)
End Sub


Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab

    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0

        If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0

    End If
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf.AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    txt.Text = Val(vsf.TextMatrix(NewRow, vsf.ColIndex("宽度")))
    
    txt.Enabled = Trim(vsf.TextMatrix(NewRow, vsf.ColIndex("AllowColResize")))

End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    '编辑处理
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub

