VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm单据选择 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "相同单号处方选择"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "Frm单据选择.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6450
      TabIndex        =   2
      Top             =   3000
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5220
      TabIndex        =   1
      Top             =   3000
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf单据选择 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Frm单据选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BlnSelect As Boolean
Dim StrNo As String
Dim IntBill As Integer
Public RecThis As New ADODB.Recordset
Private LngLastRow As Long

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    BlnSelect = True
    StrNo = Msf单据选择.TextMatrix(Msf单据选择.Row, 1)
    IntBill = Msf单据选择.TextMatrix(Msf单据选择.Row, 2)
    
    Unload Me
End Sub

Public Function ShowMe(ByVal FrmParent As Form, ByVal RecObject As ADODB.Recordset) As String
    Set RecThis = RecObject.Clone
    Me.Show 1, FrmParent
    If BlnSelect Then ShowMe = StrNo & ";" & IntBill
End Function

Private Sub Form_Activate()
    Msf单据选择.Row = 1
    Msf单据选择_EnterCell
End Sub

Private Sub Form_Load()
    BlnSelect = False
    StrNo = "": IntBill = 0
    
    With Msf单据选择
        Set .DataSource = RecThis
        .ColWidth(0) = 500
        .ColWidth(1) = 800
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 500
        .ColWidth(9) = 800
        .ColWidth(10) = 1200
    End With
End Sub

Private Sub Msf单据选择_DblClick()
    CmdOK_Click
End Sub

Private Sub Msf单据选择_EnterCell()
    Dim LngSelectRow As Long, intCol As Integer
    With Msf单据选择
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If LngLastRow > 0 And LngLastRow < .Rows Then
            .Row = LngLastRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngLastRow = LngSelectRow
        .Row = LngLastRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub
