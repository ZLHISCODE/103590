VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmParReason 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数变动原因登记"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   10470
      TabIndex        =   0
      Top             =   4395
      Width           =   10470
      Begin VB.CommandButton cmdOK 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   9120
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   7695
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfReason 
      Height          =   3885
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   10245
      _cx             =   18071
      _cy             =   6853
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   8421376
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   12
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmParReason.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "以下参数影响重大，调整参数后必须登记原因。"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmParReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mrsPar As ADODB.Recordset
Private Enum constCol
    col_序号 = 0
    col_参数名 = 1
    col_警告说明 = 2
    col_调整原因 = 3
End Enum

Public Sub ShowMe(frmParent As Form, ByRef rsPar As ADODB.Recordset)
'参数： frmParent   -父窗体
'       rsPar       -已做了过滤的变动了值的关键参数记录集
    Set mrsPar = rsPar
    Me.Show vbModal, frmParent
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim strTmp As String
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim strDate As String
    
    arrSQL = Array()
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI:SS')"
    
    With vsfReason
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 3)) = "" Then
                lblPrompt.Caption = "第" & i & "行未输入原因，要求必须输入！"
                Exit Sub
            Else    'ID:变动内容(原值-->新值):调整原因
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                
                '    参数id_In     Zlparachangedlog.参数id%Type,
                '    变动内容_In   Zlparachangedlog.变动内容%Type, --原值-->新值
                '    变动原因_In   Zlparachangedlog.变动原因%Type,
                '    操作员姓名_In Zlparachangedlog.变动人%Type,
                '    变动时间_In   Zlparachangedlog.变动时间%Type
                strTmp = .Cell(flexcpData, i, col_参数名) & ",'" & .Cell(flexcpData, i, col_调整原因) & "','" & _
                        .TextMatrix(i, col_调整原因) & "','" & gstrUserName & "'," & strDate
                
                arrSQL(UBound(arrSQL)) = "zl_Parameters_Change_Value(" & strTmp & ")"
            End If
        Next
    End With
    strTmp = Mid(strTmp, 2)
    
    On Error GoTo ErrHandle
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    mblnOk = True
    Unload Me
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    mblnOk = False
    With vsfReason
        .Rows = mrsPar.RecordCount + 1
        
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, col_序号) = i
            .TextMatrix(i, col_参数名) = mrsPar!参数名
            .Cell(flexcpData, i, col_参数名) = Val(mrsPar!ID)
            .TextMatrix(i, col_警告说明) = mrsPar!警告说明
            .Cell(flexcpData, i, col_调整原因) = mrsPar!参数值 & "-->" & mrsPar!参数新值
            
            mrsPar.MoveNext
        Next
        .AutoSize col_警告说明
        
        .Select .FixedRows, 0, .Rows - 1, .Cols - 2
        .FillStyle = flexFillRepeat
        .CellBackColor = &HFDFCF5   '&HFAFAFA      '浅蓝(不可输入列)
        .FillStyle = flexFillSingle
        
        .Select 1, col_调整原因
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnOk = False Then Cancel = True
End Sub

Private Sub vsfReason_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col_调整原因 Then Cancel = True
End Sub

Private Sub vsfReason_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If vsfReason.Row < vsfReason.Rows - 1 Then
            vsfReason.Row = vsfReason.Row + 1
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub vsfReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("|") Then KeyAscii = 0
End Sub
