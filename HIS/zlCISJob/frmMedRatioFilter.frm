VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMedRatioFilter 
   Caption         =   "附加条件"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7905
   Icon            =   "frmMedRatioFilter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   7905
   StartUpPosition =   1  '所有者中心
   Begin RichTextLib.RichTextBox rtfInfo 
      Height          =   3225
      Left            =   165
      TabIndex        =   4
      Top             =   1065
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   5689
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMedRatioFilter.frx":6852
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   300
      Left            =   5580
      TabIndex        =   2
      Top             =   4575
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   6735
      TabIndex        =   1
      Top             =   4575
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl tbcFilter 
      Height          =   3795
      Left            =   30
      TabIndex        =   0
      Top             =   615
      Width           =   7665
      _Version        =   589884
      _ExtentX        =   13520
      _ExtentY        =   6694
      _StockProps     =   64
   End
   Begin VB.Line lin1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   4575
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line lin2 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   7620
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmMedRatioFilter.frx":68EF
      Height          =   555
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   7710
   End
End
Attribute VB_Name = "frmMedRatioFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsFilter As ADODB.Recordset

Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objItem As TabControlItem
    Dim objPane As Pane
    Dim StrSQL As String
    
    StrSQL = "Select 类别,内容 From 药比附加条件"
    
    On Error GoTo errH
    Set mrsFilter = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    With Me.tbcFilter
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
        End With
        Set objItem = .InsertItem(0, "分类统计", rtfInfo.hwnd, 0): objItem.Color = 15790320
        Set objItem = .InsertItem(1, "抗菌药物", rtfInfo.hwnd, 0): objItem.Color = &HC0C0FF
        Set objItem = .InsertItem(2, "基本药物", rtfInfo.hwnd, 0): objItem.Color = &HAC0FF
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'保存值然后卸载窗体
    Dim strTmp As String
    
    Select Case tbcFilter.Selected.Index
        Case 0
            strTmp = "'分类统计'"
        Case 1
            strTmp = "'抗菌药物'"
        Case 2
            strTmp = "'基本药物'"
    End Select
    
    strTmp = "zl_药比附加条件_update(" & strTmp & ",'" & rtfInfo.Text & "')"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strTmp, Me.Caption)
    Unload Me
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lblInfo.Top = 80
    lblInfo.Left = 70
    lblInfo.Width = Me.ScaleWidth - 70
    
    cmdCancel.Top = Me.ScaleHeight - cmdCancel.Height - 40
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 20
    
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 10
    
    lin1.y1 = cmdCancel.Top - 50
    lin1.y2 = lin1.y1
    
    lin2.y1 = cmdCancel.Top - 45
    lin2.y2 = lin2.y1
    
    lin1.x2 = Me.Width
    lin2.x2 = Me.Width
    
    tbcFilter.Top = lblInfo.Height + 10
    tbcFilter.Left = 30
    tbcFilter.Width = lblInfo.Width - 10
    tbcFilter.Height = cmdOK.Top - 650
    
End Sub

Private Sub tbcFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mrsFilter Is Nothing Then Exit Sub
    If mrsFilter.EOF Then Exit Sub
    Select Case Item.Index
        Case 0
            mrsFilter.Filter = "类别='分类统计'"
            rtfInfo.Text = "" & mrsFilter!内容
        Case 1
            mrsFilter.Filter = "类别='抗菌药物'"
            rtfInfo.Text = "" & mrsFilter!内容
        Case 2
            mrsFilter.Filter = "类别='基本药物'"
            rtfInfo.Text = "" & mrsFilter!内容
    End Select
End Sub
