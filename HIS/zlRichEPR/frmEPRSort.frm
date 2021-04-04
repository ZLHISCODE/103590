VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRSort 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病历文件排序"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   Icon            =   "frmEPRSort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkShare 
      Caption         =   "显示非共享页面文件"
      Height          =   270
      Left            =   3480
      TabIndex        =   4
      Top             =   3795
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   6870
      TabIndex        =   2
      Top             =   3735
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   5610
      TabIndex        =   1
      Top             =   3735
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      DragIcon        =   "frmEPRSort.frx":6852
      Height          =   3630
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   7950
      _cx             =   14023
      _cy             =   6403
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
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   45
      Picture         =   "frmEPRSort.frx":D0A4
      Top             =   3750
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "鼠标左键拖放来改变排列顺序。"
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   3
      Top             =   3795
      Width           =   2610
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEPRSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'窗体常量
'-----------------------------------------------------
Private Enum mCol
    序号 = 0: ID = 1
End Enum
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng种类 As Long
Private mstr页面编号 As String
Private mblnOk As Boolean
Private Sub FillList()
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

    If chkShare.Value = vbUnchecked Then
        gstrSQL = "Select r.序号, r.Id, r.病历名称, r.创建人 As 创建人, To_Char(r.创建时间, 'yyyy-mm-dd hh24:mi') As 创建时间," & vbNewLine & _
                    "       To_Char(r.完成时间, 'yyyy-mm-dd hh24:mi') As 完成时间, r.最后版本 As 版本" & vbNewLine & _
                    "From 电子病历记录 R, (Select ID From 病历文件列表 Where 页面 = [4]) F" & vbNewLine & _
                    "Where r.文件id = f.Id And r.病人来源 = 2 And r.病历种类 = [3] And r.病人id = [1] And r.主页id = [2]" & vbNewLine & _
                    "Order By r.序号, r.创建时间"
    Else
        gstrSQL = "Select r.序号, r.Id, r.病历名称, r.创建人 As 创建人, To_Char(r.创建时间, 'yyyy-mm-dd hh24:mi') As 创建时间," & vbNewLine & _
                    "       To_Char(r.完成时间, 'yyyy-mm-dd hh24:mi') As 完成时间, r.最后版本 As 版本" & vbNewLine & _
                    "From 电子病历记录 R" & vbNewLine & _
                    "Where r.病人来源 = 2 And r.病历种类 = [3] And r.病人id = [1] And r.主页id = [2]" & vbNewLine & _
                    "Order By r.序号, r.创建时间"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID, mlng种类, mstr页面编号)

    
    With Me.vfgThis
        .Clear
        .FixedCols = 0
        Set .DataSource = rsTemp
        .FixedCols = 1: .ColWidth(mCol.ID) = 0: .ColHidden(mCol.ID) = True
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        .ColAlignment(mCol.序号) = flexAlignCenterCenter
        For lngCount = .FixedRows To .Rows - 1
            .TextMatrix(lngCount, mCol.序号) = lngCount
        Next
    End With
End Sub
Public Function ShowMe(ByRef frmParent As Object, _
    Optional ByVal lng病人ID As Long, _
    Optional ByVal lng主页ID As Long, _
    Optional ByVal lng种类 As Long, _
    Optional ByVal str页面编号 As String) As Boolean
    
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng种类 = lng种类
    mstr页面编号 = str页面编号
    
    Call FillList
    Me.Show vbModal, frmParent
    ShowMe = mblnOk
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkShare_Click()
    Call FillList
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Err = 0: On Error GoTo LL
    Dim i As Long

    For i = 1 To vfgThis.Rows - 1
        gstrSQL = "Zl_电子病历记录_更改序号(" & Val(vfgThis.Cell(flexcpText, i, mCol.ID)) & "," & Val(vfgThis.Cell(flexcpText, i, mCol.序号)) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    mblnOk = True
    Unload Me
    Exit Sub
LL:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    mblnOk = False
End Sub
Private Sub vfgThis_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If vfgThis.ROW < vfgThis.Rows - 1 Then vfgThis.ROW = vfgThis.ROW + 1
    End If
End Sub

Private Sub vfgThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '鼠标拖动
    With vfgThis
        If Button = vbLeftButton Then
'            Cancel = True
            Dim r%
            r = .ROW
            .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = vbRed
            r = .DragRow(r)
            .Cell(flexcpCustomFormat, r, 0, r, .Cols - 1) = False
            Dim i As Long
            For i = 1 To vfgThis.Rows - 1
                vfgThis.Cell(flexcpText, i, mCol.序号) = i
            Next
        End If
    End With
End Sub




