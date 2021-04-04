VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugSumAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "查询条件设置"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraRangeSelect 
      Caption         =   "范围选择"
      Height          =   1530
      Left            =   105
      TabIndex        =   3
      Top             =   150
      Width           =   2910
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   990
         TabIndex        =   2
         Top             =   1035
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   61472771
         CurrentDate     =   36257
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   990
         TabIndex        =   1
         Top             =   690
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   61472771
         CurrentDate     =   36257
      End
      Begin VB.Label lbl库房 
         Alignment       =   1  'Right Justify
         Caption         =   "库房"
         Height          =   180
         Left            =   285
         TabIndex        =   6
         Top             =   390
         Width           =   660
      End
      Begin VB.Label lblStartDate 
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期"
         Height          =   180
         Left            =   195
         TabIndex        =   5
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblEndDate 
         BackStyle       =   0  'Transparent
         Caption         =   "终止日期"
         Height          =   180
         Left            =   210
         TabIndex        =   4
         Top             =   1095
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3390
      TabIndex        =   8
      Top             =   555
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3390
      TabIndex        =   7
      Top             =   150
      Width           =   1100
   End
End
Attribute VB_Name = "frmDrugSumAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnAskOk As Boolean
Dim rsRoom As New ADODB.Recordset
Public inDeptId As Long
Dim blnFirst As Boolean
Private Sub CmdCancel_Click()
    blnAskOk = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    blnAskOk = True
    Me.Hide
End Sub

Private Sub dtpStartDate_Change()
    If Me.dtpStartDate.Value > Me.dtpEndDate.Value Then
        Me.dtpEndDate.Value = Me.dtpStartDate.Value
    End If
End Sub

Private Sub dtpEndDate_Change()
    If Me.dtpStartDate.Value > Me.dtpEndDate.Value Then
        Me.dtpStartDate.Value = Me.dtpEndDate.Value
    End If
End Sub

Private Sub Form_Activate()
    Dim iRow As Long
    If Not blnFirst Then Exit Sub
'    For iRow = 0 To cbo库房.ListCount - 1
'        If Me.cbo库房.ItemData(iRow) = inDeptId Then
'            Me.cbo库房.ListIndex = iRow
'            Exit For
'        End If
'    Next
    Dim i As Integer
    
    cbo库房.Clear
    With frmDrugQuery.cob库房
         For i = 0 To .ListCount - 1
            cbo库房.AddItem .List(i)
            cbo库房.ItemData(cbo库房.NewIndex) = .ItemData(i)
            If .ItemData(i) = inDeptId Then
                cbo库房.ListIndex = cbo库房.NewIndex
            End If
         Next
    End With
    
    If InStr(gstrStockSearchPrivs, "所有库房") = 0 Then Me.cbo库房.Enabled = False Else Me.cbo库房.Enabled = True
    
End Sub

Private Sub Form_Load()
    Dim StrSql As String
    blnFirst = True
    
'    With rsRoom
'        If .State = adStateOpen Then .Close
'        StrSql = "Select distinct a.ID,a.编码,a.名称 From 部门表 a,部门性质说明 b,部门性质分类 C " & _
'                " Where a.id=b.部门id And b.工作性质=c.名称 And (c.编码 in ('H','I','J','K','L','M','N')) "
'        rsRoom.Open StrSql, gcnOracle, adOpenKeyset
'        Me.cbo库房.Clear
'        Me.cbo库房.AddItem "所有库房"
'        Me.cbo库房.ItemData(Me.cbo库房.NewIndex) = 0
'        Me.cbo库房.ListIndex = Me.cbo库房.NewIndex
'        Do While Not .EOF
'            Me.cbo库房.AddItem .Fields("名称").Value
'            Me.cbo库房.ItemData(Me.cbo库房.NewIndex) = .Fields("ID").Value
'            .MoveNext
'        Loop
'    End With
    Me.dtpEndDate.MaxDate = Currentdate()
    Me.dtpEndDate.Value = dtpEndDate.MaxDate
    Me.dtpStartDate.MaxDate = Me.dtpEndDate.Value
    Me.dtpStartDate.Value = DateAdd("m", -1, Me.dtpEndDate.Value)
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Me.MousePointer = 0
End Sub
