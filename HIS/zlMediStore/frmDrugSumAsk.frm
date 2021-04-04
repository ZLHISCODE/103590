VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugSumAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ѯ��������"
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
   StartUpPosition =   1  '����������
   Begin VB.Frame fraRangeSelect 
      Caption         =   "��Χѡ��"
      Height          =   1530
      Left            =   105
      TabIndex        =   3
      Top             =   150
      Width           =   2910
      Begin VB.ComboBox cbo�ⷿ 
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
         CustomFormat    =   "yyyy��MM��dd��"
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
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   61472771
         CurrentDate     =   36257
      End
      Begin VB.Label lbl�ⷿ 
         Alignment       =   1  'Right Justify
         Caption         =   "�ⷿ"
         Height          =   180
         Left            =   285
         TabIndex        =   6
         Top             =   390
         Width           =   660
      End
      Begin VB.Label lblStartDate 
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����"
         Height          =   180
         Left            =   195
         TabIndex        =   5
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblEndDate 
         BackStyle       =   0  'Transparent
         Caption         =   "��ֹ����"
         Height          =   180
         Left            =   210
         TabIndex        =   4
         Top             =   1095
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3390
      TabIndex        =   8
      Top             =   555
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
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
'    For iRow = 0 To cbo�ⷿ.ListCount - 1
'        If Me.cbo�ⷿ.ItemData(iRow) = inDeptId Then
'            Me.cbo�ⷿ.ListIndex = iRow
'            Exit For
'        End If
'    Next
    Dim i As Integer
    
    cbo�ⷿ.Clear
    With frmDrugQuery.cob�ⷿ
         For i = 0 To .ListCount - 1
            cbo�ⷿ.AddItem .List(i)
            cbo�ⷿ.ItemData(cbo�ⷿ.NewIndex) = .ItemData(i)
            If .ItemData(i) = inDeptId Then
                cbo�ⷿ.ListIndex = cbo�ⷿ.NewIndex
            End If
         Next
    End With
    
    If InStr(gstrStockSearchPrivs, "���пⷿ") = 0 Then Me.cbo�ⷿ.Enabled = False Else Me.cbo�ⷿ.Enabled = True
    
End Sub

Private Sub Form_Load()
    Dim StrSql As String
    blnFirst = True
    
'    With rsRoom
'        If .State = adStateOpen Then .Close
'        StrSql = "Select distinct a.ID,a.����,a.���� From ���ű� a,��������˵�� b,�������ʷ��� C " & _
'                " Where a.id=b.����id And b.��������=c.���� And (c.���� in ('H','I','J','K','L','M','N')) "
'        rsRoom.Open StrSql, gcnOracle, adOpenKeyset
'        Me.cbo�ⷿ.Clear
'        Me.cbo�ⷿ.AddItem "���пⷿ"
'        Me.cbo�ⷿ.ItemData(Me.cbo�ⷿ.NewIndex) = 0
'        Me.cbo�ⷿ.ListIndex = Me.cbo�ⷿ.NewIndex
'        Do While Not .EOF
'            Me.cbo�ⷿ.AddItem .Fields("����").Value
'            Me.cbo�ⷿ.ItemData(Me.cbo�ⷿ.NewIndex) = .Fields("ID").Value
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
