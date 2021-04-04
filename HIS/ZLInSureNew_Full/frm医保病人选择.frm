VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm医保病人选择 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病人选择"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7875
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6255
      TabIndex        =   3
      Top             =   4395
      Width           =   1150
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4875
      TabIndex        =   2
      Top             =   4395
      Width           =   1150
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   4245
      Left            =   2670
      TabIndex        =   1
      Top             =   15
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   7488
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDept 
      Height          =   4245
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   7488
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm医保病人选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mlng病人ID  As Long

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mshPati.Rows > 1 And mshPati.TextMatrix(1, 0) <> "" Then
        mlng病人ID = mshPati.TextMatrix(mshPati.Row, 0)
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub mshDept_EnterCell()
    Dim i As Integer
    Dim rsPati As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Me.Refresh
    mshPati.Clear
    If mshDept.RowData(mshDept.Row) = 0 Then Exit Sub
    
    '当前在院病人
    gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
    If Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34 Then
        gstrSQL = "Select A.病人ID,A.住院号,A.姓名,A.当前床号 as 床位,A.性别,A.费别" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID And Nvl(A.主页ID,0)=C.主页ID " & _
            "       and A.当前科室ID = " & mshDept.RowData(mshDept.Row) & _
            "       and C.险类 is null and A.在院=1" & _
            " Order by A.住院号 Desc"
    Else
        gstrSQL = "Select A.病人ID,A.住院号,A.姓名,A.当前床号 as 床位,A.性别,A.费别" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID And Nvl(A.住院次数,0)=C.主页ID " & _
            "       and A.当前科室ID = " & mshDept.RowData(mshDept.Row) & _
            "       and C.险类 is null and A.在院=1" & _
            " Order by A.住院号 Desc"
    End If
    Call OpenRecordset(rsPati, Me.Caption)
    
    With rsPati
        If Not .EOF Then
            Set mshPati.Recordset = rsPati
            mshPati.ColWidth(0) = 800
            mshPati.ColWidth(1) = 800
            mshPati.ColWidth(2) = 850
            mshPati.ColWidth(3) = 600
            mshPati.ColWidth(4) = 500
            mshPati.ColWidth(5) = 800
            mshPati.ColAlignment(4) = 4
            mshPati.ColAlignment(5) = 1
        Else
            mshPati.Rows = 2
            mshPati.Cols = 2
        End If
    End With
    
    For i = 0 To mshPati.Cols - 1
        mshPati.ColAlignmentFixed(i) = 4
    Next
    mshPati.RowHeight(0) = 320
    mshPati.Row = 1: mshPati.TopRow = 1
    mshPati.COL = 0: mshPati.ColSel = mshPati.Cols - 1
    
    If Not rsPati.EOF Then
        If Visible Then mshPati.SetFocus
    Else
        If Visible Then mshDept.SetFocus
    End If
End Sub

Private Sub mshDept_GotFocus()
    mshDept.BackColorSel = &H8000000D
    mshPati.BackColorSel = &H808080
End Sub

Private Sub mshDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mshPati_KeyDown(KeyCode, Shift)
End Sub

Private Sub mshPati_DblClick()
    cmdOK_Click
End Sub

Private Sub mshPati_GotFocus()
    mshDept.BackColorSel = &H808080
    mshPati.BackColorSel = &H8000000D
End Sub

Private Sub mshPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub Form_Activate()
    mshPati.SetFocus
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
        
    mshDept.TextMatrix(0, 0) = "编码"
    mshDept.TextMatrix(0, 1) = "科室"
    mshDept.Rows = 2: mshDept.Cols = 2
    mshDept.ColAlignmentFixed(0) = 4
    mshDept.ColAlignmentFixed(1) = 4
    mshDept.ColAlignment(0) = 1
    mshDept.ColAlignment(1) = 1
    mshDept.ColWidth(0) = 830
    mshDept.ColWidth(1) = 1500
    mshDept.Row = 1
    
    gstrSQL = "Select Distinct D.ID,D.编码,D.名称 " & _
        " From 部门表 D,部门性质说明 N " & _
        " Where D.ID=N.部门ID and N.工作性质 IN('临床','手术') and N.服务对象 IN (2,3)" & _
        " And (D.撤档时间 is NULL or D.撤档时间=TO_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by D.编码"
    
    rsTmp.CursorLocation = adUseClient
    Call OpenRecordset(rsTmp, Me.Caption)
        
    With rsTmp
        If Not .EOF Then
            mshDept.Rows = rsTmp.RecordCount + 1
            For i = 1 To .RecordCount
                mshDept.TextMatrix(i, 0) = !编码
                mshDept.TextMatrix(i, 1) = !名称
                mshDept.RowData(i) = !ID
                
                .MoveNext
            Next
        End If
    End With
    
    mshDept.COL = 0: mshDept.ColSel = mshDept.Cols - 1
    mshDept.TopRow = mshDept.Row
    Call mshDept_EnterCell
End Sub

Private Sub mshPati_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        If mshDept.RowData(mshDept.Row) = 0 Then Exit Sub
        If KeyCode = vbKeyLeft Then
            If mshDept.Row - 1 >= 1 Then mshDept.Row = mshDept.Row - 1
        ElseIf KeyCode = vbKeyRight Then
            If mshDept.Row + 1 <= mshDept.Rows - 1 Then
                mshDept.Row = mshDept.Row + 1
            End If
        End If
        mshDept.COL = 0: mshDept.ColSel = mshDept.Cols - 1
        If mshDept.CellTop + mshDept.CellHeight > mshDept.Height - 300 Then mshDept.TopRow = mshDept.TopRow + 1
        Call mshDept_EnterCell
        mshPati.COL = 0: mshPati.ColSel = mshPati.Cols - 1
    End If
End Sub

Public Function Get病人(lng病人ID As Long) As Boolean
'得到病人
    mblnOK = False
    
    frm医保病人选择.Show vbModal
    
    If mblnOK = True Then
        lng病人ID = mlng病人ID
    End If
    Get病人 = mblnOK
End Function
