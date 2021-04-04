VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAppChkRpt 
   Caption         =   "系统检查报告"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "frmAppChkRpt.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10215
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   8910
      TabIndex        =   2
      Top             =   6660
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   7725
      TabIndex        =   1
      Top             =   6660
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdReport 
      Height          =   6480
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   11430
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      FillStyle       =   1
      MergeCells      =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label lblWarn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注意:存在严重或较重的问题,请仔细核查!"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   6795
      Visible         =   0   'False
      Width           =   3330
   End
End
Attribute VB_Name = "frmAppChkRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnModifyCheck As Boolean  '修正对象检查
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    objPrint.Title.Text = "系统检查报告"
    objRow.Add "打印时间：" & Format(CurrentDate, "yyyy年MM月dd日")
    Set objPrint.Body = hgdReport
    objPrint.BelowAppRows.Add objRow
    Select Case zlPrintAsk(objPrint)
        Case 1
             zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
    End Select
End Sub

Private Sub Form_Activate()
    hgdReport.Redraw = True
End Sub
Private Sub Form_Load()
    Call InitGrid
End Sub
Private Sub InitGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格列头信息
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-08-26 14:21:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With hgdReport
        .Redraw = False
        If mblnModifyCheck = False Then
             .Rows = 1: .Clear
            .Cols = 3
            .TextMatrix(0, 0) = "对象"
            .TextMatrix(0, 1) = "检查情况"
            .TextMatrix(0, 2) = "影响程度"
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .ColAlignmentFixed(0) = 4
            .ColAlignmentFixed(1) = 4
            .ColAlignmentFixed(2) = 4
            .ColWidth(0) = 2200
            .ColWidth(1) = 3000
            .ColWidth(2) = 3000
            .ColData(0) = 2200
            .ColData(1) = 3000
            .ColData(2) = 3000
            .MergeCol(0) = True
            Exit Sub
        End If
        i = 0: .Cols = 6: .Rows = 1: .Clear
        .TextMatrix(0, i) = "类型": .ColAlignment(i) = 1: .ColAlignmentFixed(i) = 4: i = i + 1
        .TextMatrix(0, i) = "对象名": .ColAlignment(i) = 1: .ColAlignmentFixed(i) = 4: i = i + 1
        .TextMatrix(0, i) = "错误类型": .ColAlignment(i) = 1: .ColAlignmentFixed(i) = 4: i = i + 1
        .TextMatrix(0, i) = "错误影响程度": .ColAlignment(i) = 1: .ColAlignmentFixed(i) = 4: i = i + 1
        
        .TextMatrix(0, i) = "修正情况": .ColAlignment(i) = 1: .ColAlignmentFixed(i) = 4: i = i + 1
        .TextMatrix(0, i) = "修正说明:": .ColAlignment(i) = 1: .ColAlignmentFixed(i) = 4: i = i + 1
        For i = 0 To .Cols - 1
            .ColWidth(i) = 1000: .ColData(i) = .ColWidth(i)
        Next
    End With
End Sub
Private Sub Form_Resize()
    Dim i As Long, sngColWidthTemp As Single
    On Error Resume Next
    Dim sngColWidth As Single
    With hgdReport
        .Top = ScaleTop
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Height = ScaleHeight - cmdPrint.Height - 240
        For i = 0 To .Cols - 1
            sngColWidth = sngColWidth + .ColData(i)
            If i < .Cols - 1 Then sngColWidthTemp = sngColWidthTemp + .ColWidth(i)
        Next
        If sngColWidth <> 0 And sngColWidth < .Width Then
            For i = 0 To .Cols - 2
                .ColWidth(i) = .ColData(i)
            Next
            .ColWidth(.Cols - 1) = IIf(.Width - sngColWidthTemp - 320 < 0, 100, sngColWidthTemp)
        Else
            If .Width - sngColWidthTemp - 320 < 0 Then
                If mblnModifyCheck Then
                Else
                  .ColWidth(2) = 3000
                  .ColWidth(1) = .Width - .ColWidth(0) - .ColWidth(2) - 320
                End If
            Else
                If mblnModifyCheck Then
                Else
                    .ColWidth(1) = .Width - .ColWidth(0) - .ColWidth(2) - 320
                End If
            End If
        End If
    End With
    
    cmdClose.Top = hgdReport.Top + hgdReport.Height + 120
    cmdClose.Left = ScaleWidth - cmdClose.Width - 120
    cmdPrint.Top = cmdClose.Top
    cmdPrint.Left = cmdClose.Left - cmdPrint.Width - 120
    
    lblWarn.Top = cmdClose.Top + (cmdClose.Height - lblWarn.Height) / 2
End Sub


Public Property Get blnModiyfyCheck() As Boolean
   blnModiyfyCheck = mblnModifyCheck
End Property

Public Property Let blnModiyfyCheck(ByVal vNewValue As Boolean)
    mblnModifyCheck = vNewValue
    If vNewValue Then
        Me.Caption = "系统修正报告"
    Else
        Me.Caption = "系统检查报告"
    End If
    Call InitGrid
    
End Property
