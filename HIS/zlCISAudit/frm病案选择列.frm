VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm病案选择列 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "列选择器"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4920
   Icon            =   "frm病案选择列.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VSFlex8Ctl.VSFlexGrid vfgColumn 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _cx             =   6165
      _cy             =   6376
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3630
      TabIndex        =   10
      Top             =   90
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   9
      Top             =   510
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3630
      TabIndex        =   8
      Top             =   1290
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtWidth 
      Height          =   300
      Left            =   1170
      TabIndex        =   1
      Top             =   3810
      Width           =   1155
   End
   Begin VB.ComboBox cmbAlign 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4230
      Width           =   1185
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全选(&S)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3630
      TabIndex        =   7
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全清(&L)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3630
      TabIndex        =   6
      Top             =   2460
      Width           =   1100
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "上移(&U)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3630
      TabIndex        =   5
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "下移(&D)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3630
      TabIndex        =   4
      Top             =   3360
      Width           =   1100
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复默认设置(&R)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3030
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label lblAlign 
      Caption         =   "对齐方式(&A)"
      Height          =   180
      Left            =   60
      TabIndex        =   12
      Top             =   4290
      Width           =   990
   End
   Begin VB.Label lblWidth 
      Caption         =   "列宽(&W)"
      Height          =   180
      Left            =   420
      TabIndex        =   11
      Top             =   3870
      Width           =   630
   End
End
Attribute VB_Name = "frm病案选择列"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mVsGrid As VSFlexGrid
Private mblnOK As Boolean

Public Function ShowColSet(ByVal frmMain As Form, ByVal strTittle As String, vsGrid As VSFlexGrid) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:列设置接口
    '参数:
    '返回:列设置成功,返回true,否则返回False
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Err = 0: On Error Resume Next
    Set mVsGrid = vsGrid
    If strTittle <> "" Then Me.Caption = strTittle
    
    cmbAlign.AddItem "左上对齐"
    cmbAlign.AddItem "左中对齐"
    cmbAlign.AddItem "左下对齐"
    cmbAlign.AddItem "中上对齐"
    cmbAlign.AddItem "居中对齐"
    cmbAlign.AddItem "中下对齐"
    cmbAlign.AddItem "右上对齐"
    cmbAlign.AddItem "右中对齐"
    cmbAlign.AddItem "右下对齐"
    
    Call LoadFulltoColSel
'    Call cmdRestore_Click
    With Me
        .Show vbModal, frmMain
    End With
    ShowColSet = mblnOK
End Function

Private Function LoadFulltoColSel() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载列设置
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:lesfeng
    '日期:2009-08-25 16:46:43
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, lngRow As Long, arrSplit As Variant
    Dim sngFrmHeight As Single, sngSelSumHeight As Single

    Call initVfgColumnTitle
    With mVsGrid
        lngRow = 1
        For i = 0 To .Cols - 1
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            arrSplit = Split(.ColData(i) & "||", "||")
            
            If Trim(.ColKey(i)) <> "" And (Val(arrSplit(0)) = 1 Or Val(arrSplit(0)) = 0) Then
                vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("列名")) = .ColKey(i)
                vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("选择")) = IIf(.ColWidth(i) = 0 Or .ColHidden(i), False, True)
                vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("对齐")) = .ColAlignment(i)
                vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("列宽")) = .ColWidth(i)
                vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("固定")) = arrSplit(0)
                If .ColWidth(i) = 0 Or .ColHidden(i) Then
                    vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("原值")) = 0
                    vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("改变")) = 0
                Else
                    vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("原值")) = 1
                    vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("改变")) = 1
                End If
                If Val(arrSplit(0)) = 1 Then
                    vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("备注")) = "不能隐藏"
                End If

                vfgColumn.RowData(lngRow) = Val(arrSplit(0))
                If Val(arrSplit(0)) = 1 Then
                    vfgColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, vfgColumn.Cols - 1) = vbBlue
                End If
                vfgColumn.Rows = vfgColumn.Rows + 1
                lngRow = lngRow + 1
            End If
        Next
    End With
    
    If vfgColumn.Rows > 2 Then vfgColumn.Rows = vfgColumn.Rows - 1
    With vfgColumn
        .Visible = True
        .Editable = flexEDKbdMouse
        .ZOrder 0
        .ColDataType(.ColIndex("选择")) = flexDTBoolean
'        '列排序并拖动
'        .ExplorerBar = flexExSortShowAndMove
        '行选择
        .SelectionMode = flexSelectionByRow

        If .Rows > 1 Then
            .Row = 1
            .Select .Row, .ColIndex("选择")
            Call vfgColumn_Click
            Call setenabled
        End If
        .SetFocus
    End With
    Call SetcmbAlign
End Function

Private Function SetVsGridCol(ByVal strColKey As String, ByVal blnShow As Boolean, ByVal blnBatch As Boolean, ByVal lngColWidth As Long, ByVal lngAlign As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置显示列
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:lesfeng
    '日期:2009-08-25 16:46:43
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, lngRow As Long
        
    With mVsGrid
        .ColHidden(.ColIndex(strColKey)) = Not blnShow
        If .ColWidth(.ColIndex(strColKey)) = 0 Then .ColWidth(.ColIndex(strColKey)) = 1000
        If lngColWidth >= 0 Then .ColWidth(.ColIndex(strColKey)) = lngColWidth
        .ColAlignment(.ColIndex(strColKey)) = lngAlign
        '问题29530 by lesfeng 2010-05-06
        If .Rows > 1 Then .Cell(flexcpAlignment, .FixedRows, .ColIndex(strColKey), .Rows - 1, .ColIndex(strColKey)) = lngAlign
    End With
End Function

Private Function SaveData() As Boolean
    Dim strOldValue As String
    Dim strNewValue As String
    Dim lngColWith As Long
    Dim lngAlign As Long
    Dim lngRow As Long
    Dim blnShow As Boolean
    
    With vfgColumn
        For lngRow = 1 To .Rows - 1
            strOldValue = .TextMatrix(lngRow, .ColIndex("原值"))
            strNewValue = .TextMatrix(lngRow, .ColIndex("改变"))
            If strOldValue <> strNewValue Then
                 blnShow = GetVsGridBoolColVal(vfgColumn, lngRow, .ColIndex("选择"))
                 lngColWith = Val(.TextMatrix(lngRow, .ColIndex("列宽")))
                 lngAlign = .TextMatrix(lngRow, .ColIndex("对齐"))
                 Call SetVsGridCol(.TextMatrix(lngRow, .ColIndex("列名")), blnShow, IIf(.Tag = "Head", False, True), lngColWith, lngAlign)
            End If
        Next
    End With
End Function

Private Sub cmbAlign_Click()
    Dim strAlign As String
    
    strAlign = cmbAlign.Text
    With vfgColumn
        If .Row > 0 Then
            Select Case strAlign
            Case "左上对齐"
                .TextMatrix(.Row, .ColIndex("对齐")) = 0
            Case "左中对齐"
                .TextMatrix(.Row, .ColIndex("对齐")) = 1
            Case "左下对齐"
                .TextMatrix(.Row, .ColIndex("对齐")) = 2
            Case "中上对齐"
                .TextMatrix(.Row, .ColIndex("对齐")) = 3
            Case "居中对齐"
                .TextMatrix(.Row, .ColIndex("对齐")) = 4
            Case "中下对齐"
                .TextMatrix(.Row, .ColIndex("对齐")) = 5
            Case "右上对齐"
                .TextMatrix(.Row, .ColIndex("对齐")) = 6
            Case "右中对齐"
                .TextMatrix(.Row, .ColIndex("对齐")) = 7
            Case "右下对齐"
                .TextMatrix(.Row, .ColIndex("对齐")) = 8
            End Select
            .TextMatrix(.Row, .ColIndex("改变")) = 2
        End If
    End With
End Sub

Private Sub SetcmbAlign()
    Dim strAlign As String
    Dim lngColWith As Long
    
    With vfgColumn
        If .Row > 0 Then
            strAlign = .TextMatrix(.Row, .ColIndex("对齐"))
            lngColWith = Val(.TextMatrix(.Row, .ColIndex("列宽")))
        Else
            Exit Sub
        End If
    End With

    Select Case Val(strAlign)
    Case 0
        cmbAlign.Text = "左上对齐"
    Case 1
        cmbAlign.Text = "左中对齐"
    Case 2
        cmbAlign.Text = "左下对齐"
    Case 3
        cmbAlign.Text = "中上对齐"
    Case 4
        cmbAlign.Text = "居中对齐"
    Case 5
        cmbAlign.Text = "中下对齐"
    Case 6
        cmbAlign.Text = "右上对齐"
    Case 7
        cmbAlign.Text = "右中对齐"
    Case 8
        cmbAlign.Text = "右下对齐"
    End Select
    txtWidth.Text = lngColWith
End Sub

Private Sub setenabled()
    With vfgColumn
        If .Row > 0 Then
            If .Row = 1 Then
                cmdUp.Enabled = False
            Else
                cmdUp.Enabled = True
            End If
            If .Row = .Rows - 1 Then
                cmdDown.Enabled = False
            Else
                cmdDown.Enabled = True
            End If
        Else
            cmdUp.Enabled = False
            cmdDown.Enabled = False
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    mblnOK = False
End Sub

Private Sub cmdClear_Click()
    Dim lngRow As Long
    
    With vfgColumn
        For lngRow = 1 To .Rows - 1
             If Val(.RowData(lngRow)) = 0 Then
                If .TextMatrix(lngRow, .ColIndex("选择")) = "0" Then
                Else
                    .TextMatrix(lngRow, .ColIndex("改变")) = 0
                    .TextMatrix(lngRow, .ColIndex("选择")) = False
                End If
            End If
        Next
    End With
End Sub

Private Sub CmdDown_Click()
    With vfgColumn
        If .Row = .Rows - 1 Then
        Else
            .Select .Row + 1, .ColIndex("选择")
            Call vfgColumn_Click
        End If
        Call setenabled
    End With
End Sub

Private Sub cmdOK_Click()
    Call SaveData
    Unload Me
    mblnOK = True
End Sub

Private Sub cmdRestore_Click()
    Call LoadFulltoColSel
End Sub

Private Sub cmdSelect_Click()
    Dim lngRow As Long
    
    With vfgColumn
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("选择")) Then
            Else
                .TextMatrix(lngRow, .ColIndex("改变")) = 1
                .TextMatrix(lngRow, .ColIndex("选择")) = True
            End If
        Next
    End With
End Sub

Private Sub CmdUP_Click()
    With vfgColumn
        If .Row = 1 Then
        Else
            .Select .Row - 1, .ColIndex("选择")
            Call vfgColumn_Click
        End If
        Call setenabled
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnOK = False
End Sub

Private Sub txtWidth_Change()
    If Trim(txtWidth) <> "" Then Call IsValid
End Sub

Private Function IsValid() As Boolean
    Dim blnValid As Boolean
    
    blnValid = True
    If IsNumeric(txtWidth) = False Then
        blnValid = False
        MsgBox "请输入一个合法的数值。", vbInformation, gstrSysName
    Else
        If Val(txtWidth.Text) > 10000 Or Val(txtWidth.Text) < 0 Then
            MsgBox "请输入一个小于10000的正数。", vbInformation, gstrSysName
            blnValid = False
        End If
    End If
    IsValid = blnValid
End Function

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyReturn And _
        KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> Asc(".") Then KeyAscii = 0
End Sub

Private Sub txtWidth_Validate(Cancel As Boolean)
    Cancel = Not IsValid
    If Cancel = False Then
        With vfgColumn
            .TextMatrix(.Row, .ColIndex("列宽")) = txtWidth
            .TextMatrix(.Row, .ColIndex("改变")) = 2
        End With
    End If
End Sub

Private Sub vfgColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '修改后
    Dim strColKey As String, blnShow As Boolean
    With vfgColumn
        Select Case Col
        Case .ColIndex("选择")
            blnShow = GetVsGridBoolColVal(vfgColumn, Row, .ColIndex("选择"))
            If blnShow Then
                 .TextMatrix(Row, .ColIndex("改变")) = 1
            Else
                 .TextMatrix(Row, .ColIndex("改变")) = 0
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vfgColumn_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgColumn
        Select Case Col
        Case .ColIndex("选择")
            'rowdata(i):1-固定,-1-不能选,0-可选
            If Val(.RowData(Row)) = 1 Then
                Cancel = True
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub initVfgColumnTitle()
    Dim strHead As String
    strHead = "选择,300,1,1;列名,2000,1,1;备注,1000,1,1;列宽,0,7,-1;对齐,0,7,-1;固定,0,7,-1;原值,0,7,-1;改变,0,7,-1"
    Call SetVsFlexGridChangeHead(strHead, vfgColumn, 0)
End Sub

Private Sub vfgColumn_Click()
    Call SetcmbAlign
End Sub
