VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm������ҩҽ�����⻮�� 
   Caption         =   "������ҩҽ�����⻮��"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   Icon            =   "frm������ҩҽ�����⻮��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   8490
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6000
      TabIndex        =   1
      Top             =   4530
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7215
      TabIndex        =   2
      Top             =   4530
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid billDetail 
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "���ո��������Ҽ��л�����״̬"
      Top             =   0
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   7752
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm������ҩҽ�����⻮��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsDetail As ADODB.Recordset

Public Sub ShowEditor(rsDetail As ADODB.Recordset)
    Set mrsDetail = rsDetail
    Me.Show 1
    Set rsDetail = mrsDetail
End Sub

Private Sub billDetail_DblClick()
    Dim strText As String
    strText = billDetail.TextMatrix(billDetail.Row, 0)
    strText = IIf(strText = "��", "", "��")
    billDetail.TextMatrix(billDetail.Row, 0) = strText
End Sub

Private Sub billDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then Call billDetail_DblClick
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '�����趨�޸ļ�¼��
    With mrsDetail
        .MoveFirst
        Do While Not .EOF
            !ҽ���� = IIf(billDetail.TextMatrix(.AbsolutePosition, 0) = "", 0, 1)
            .Update
            .MoveNext
        Loop
    End With
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim intRow As Integer, intRecords As Integer
    Set billDetail.DataSource = mrsDetail
    billDetail.WordWrap = True
    billDetail.ColWidth(0) = 700
    billDetail.ColWidth(2) = 0
    billDetail.ColWidth(3) = 0
    billDetail.ColWidth(4) = 0
    billDetail.ColWidth(5) = 1500
    billDetail.ColWidth(6) = 1500
    billDetail.ColWidth(10) = 500
    billDetail.ColWidth(11) = 4800
    
    '������ͷ���뷽ʽ
    intRecords = billDetail.Cols - 1
    For intRow = 0 To intRecords
        billDetail.ColAlignmentFixed(intRow) = 4
    Next
    
    '�����и�
    intRecords = mrsDetail.RecordCount
    For intRow = 1 To intRecords
        billDetail.RowHeight(intRow) = 450
    Next
    
    billDetail.Col = 0
    billDetail.ColSel = billDetail.Cols - 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    With CmdCancel
        .Top = Me.ScaleHeight - .Height - 80
        .Left = Me.ScaleWidth - .Width - 150
    End With
    With cmdOK
        .Top = CmdCancel.Top
        .Left = CmdCancel.Left - .Width - 80
    End With
    With billDetail
        .Height = CmdCancel.Top - 80
        .Width = Me.ScaleWidth
    End With
End Sub
