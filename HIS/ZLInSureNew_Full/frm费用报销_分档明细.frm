VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.2#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm费用报销_分档明细 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "分档明细"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "frm费用报销_分档明细.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2970
      TabIndex        =   1
      Top             =   3210
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4230
      TabIndex        =   2
      Top             =   3210
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   5371
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
End
Attribute VB_Name = "frm费用报销_分档明细"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbln查阅 As Boolean
Private mrsData As New ADODB.Recordset
Private Const strFormat_金额 As String = "#####0.00;-#####0.00; ;"

Public Sub ShowME(Optional bln查阅 As Boolean = False, Optional ByVal rsData As ADODB.Recordset = Nothing)
    On Error Resume Next
    mbln查阅 = bln查阅
    Set mrsData = rsData
    Me.Show 1
End Sub

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strInput As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With Bill
        If .TxtVisible = False Then Exit Sub
        strInput = Val(.Text)
        If Not IsNumeric(strInput) Then
            MsgBox "比例中含有非法字符！", vbInformation, gstrSysName
            Cancel = True
            .TxtSetFocus
            Exit Sub
        End If
        If Val(strInput) < 0 Then
            MsgBox "比例必须大于零！", vbInformation, gstrSysName
            Cancel = True
            .TxtSetFocus
            Exit Sub
        End If
        If Val(strInput) > 100 Then
            MsgBox "比例不能大于100%！", vbInformation, gstrSysName
            Cancel = True
            .TxtSetFocus
            Exit Sub
        End If
        
        .Text = Format(.Text, strFormat_金额)
        .TextMatrix(.Row, 3) = Format(Val(.TextMatrix(.Row, 2)) * Val(.Text) / 100, strFormat_金额)
    End With
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim lngRow As Long
    Dim cur统筹支付 As Currency
    '将更新写入公共记录集
    
    If Not mbln查阅 Then
        Set rs分档支付_米易 = New ADODB.Recordset
        With rs分档支付_米易
            If .State = 1 Then .Close
            .Fields.Append "档次", adDouble, 10  '0:表示新增
            .Fields.Append "比例", adDouble, 18, adFldIsNullable
            .Fields.Append "名称", adLongVarChar, 100, adFldIsNullable
            .Fields.Append "进入统筹", adDouble, 18, adFldIsNullable
            .Fields.Append "统筹报销", adDouble, 18, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        
        cur统筹支付 = 0
        With Bill
            For lngRow = 1 To .Rows - 1
                rs分档支付_米易.AddNew
                rs分档支付_米易!档次 = lngRow
                rs分档支付_米易!比例 = Val(.TextMatrix(lngRow, 1))
                rs分档支付_米易!名称 = .TextMatrix(lngRow, 0)
                rs分档支付_米易!进入统筹 = Val(.TextMatrix(lngRow, 2))
                rs分档支付_米易!统筹报销 = Val(.TextMatrix(lngRow, 3))
                rs分档支付_米易.Update
                cur统筹支付 = cur统筹支付 + Val(.TextMatrix(lngRow, 3))
            Next
        End With
        gComInfo_眉山.统筹支付 = cur统筹支付 + gComInfo_眉山.实际报销
        gComInfo_眉山.统筹自付 = gComInfo_眉山.进入统筹 - cur统筹支付
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsObj As New ADODB.Recordset
    On Error Resume Next
    
    If mbln查阅 Then
        Set rsObj = mrsData
    Else
        Set rsObj = rs分档支付_米易
    End If
    
    With Bill
        .ClearBill
        .Active = True
        .Rows = 1 + rsObj.RecordCount
        .Cols = 4
        .msfObj.FixedCols = 1
        
        .TextMatrix(0, 0) = "名称"
        .TextMatrix(0, 1) = "比例"
        .TextMatrix(0, 2) = "进入统筹"
        .TextMatrix(0, 3) = "报销金额"
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 800
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .msfObj.ColAlignmentFixed = 1
        .ColData(0) = 5
        .ColData(1) = 4
        .ColData(2) = 5
        .ColData(3) = 5
        
        .PrimaryCol = 1
        .LocateCol = 1
    End With
    
    With rsObj
        If .RecordCount <> 0 Then
            .Sort = "档次 asc"
            .MoveFirst
        Else
            Unload Me
            Exit Sub
        End If
        
        Do While Not .EOF
            Bill.TextMatrix(.AbsolutePosition, 0) = !名称
            Bill.TextMatrix(.AbsolutePosition, 1) = Format(!比例, strFormat_金额)
            Bill.TextMatrix(.AbsolutePosition, 2) = Format(!进入统筹, strFormat_金额)
            Bill.TextMatrix(.AbsolutePosition, 3) = Format(!统筹报销, strFormat_金额)
            .MoveNext
        Loop
    End With
    
    Bill.AllowAddRow = False
    If mbln查阅 Then
        Bill.Active = False
        cmd确定.Visible = False
        cmd取消.Caption = "确定(&O)"
    End If
End Sub
