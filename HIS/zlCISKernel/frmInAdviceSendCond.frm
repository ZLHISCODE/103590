VERSION 5.00
Begin VB.Form frmInAdviceSendCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "frmInAdviceSendCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraDetail 
      Height          =   2070
      Index           =   0
      Left            =   135
      TabIndex        =   9
      Top             =   0
      Width           =   4500
      Begin VB.CheckBox chkDrug 
         Caption         =   "��ȡҩ"
         Height          =   195
         Index           =   2
         Left            =   3225
         TabIndex        =   5
         Top             =   1695
         Value           =   1  'Checked
         Width           =   840
      End
      Begin VB.CheckBox chkDrug 
         Caption         =   "��Ժ��ҩ"
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   4
         Top             =   1695
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chkDrug 
         Caption         =   "Ժ����ҩ"
         Height          =   195
         Index           =   0
         Left            =   975
         TabIndex        =   3
         Top             =   1695
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chk�Ӱ�Ӽ� 
         Caption         =   "ִ�мӰ�Ӽ�(&V)"
         Height          =   195
         Left            =   2595
         TabIndex        =   6
         Top             =   240
         Width           =   1650
      End
      Begin VB.ListBox lstClass 
         Columns         =   4
         Height          =   1110
         ItemData        =   "frmInAdviceSendCond.frx":058A
         Left            =   195
         List            =   "frmInAdviceSendCond.frx":058C
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   495
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ҩƷ��"
         Height          =   180
         Left            =   360
         TabIndex        =   2
         Top             =   1695
         Width           =   540
      End
      Begin VB.Label lblClass 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ҫ���͵����(&T):"
         Height          =   180
         Left            =   225
         TabIndex        =   0
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3285
      TabIndex        =   8
      Top             =   2190
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2160
      TabIndex        =   7
      Top             =   2190
      Width           =   1100
   End
End
Attribute VB_Name = "frmInAdviceSendCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstr���s As String 'OUT:�������
Public mstrҩƷ As String 'OUT:Ժ����ҩ,��Ժ��ҩ,��ȡҩ
Public mblnOK As Boolean 'OUT:�Ƿ�ȷ��

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub chkDrug_Click(Index As Integer)
    If chkDrug(0).value = 0 And chkDrug(1).value = 0 And chkDrug(2).value = 0 Then
        chkDrug(Index).value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    '�������
    mstr���s = ""
    For i = 0 To lstClass.ListCount - 1
        If lstClass.Selected(i) Then
            mstr���s = mstr���s & ",'" & Chr(lstClass.ItemData(i)) & "'"
        End If
    Next
    mstr���s = Mid(mstr���s, 2)
    If mstr���s = "" Then
        MsgBox "������ѡ��һ���������", vbInformation, gstrSysName
        lstClass.SetFocus: Exit Sub
    End If
    If UBound(Split(mstr���s, ",")) + 1 = lstClass.ListCount Then
        mstr���s = ""
    End If
    
    mstrҩƷ = chkDrug(0).value & chkDrug(1).value & chkDrug(2).value
    
    gbln�Ӱ�Ӽ� = chk�Ӱ�Ӽ�.value = 1
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        j = lstClass.ListIndex
        For i = 0 To lstClass.ListCount - 1
            lstClass.Selected(i) = True
        Next
        lstClass.ListIndex = j
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        j = lstClass.ListIndex
        For i = 0 To lstClass.ListCount - 1
            lstClass.Selected(i) = False
        Next
        lstClass.ListIndex = j
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ZLCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim strPar As String
    
    mblnOK = False
    chk�Ӱ�Ӽ�.value = IIF(gbln�Ӱ�Ӽ�, 1, 0)
    
    '�������
    Call Load�������
    
    'ҩƷ���
    strPar = zlDatabase.GetPara("סԺ����ҩƷ�������", glngSys, pסԺҽ���´�, , Array(chkDrug(0), chkDrug(1), chkDrug(2)))
    chkDrug(0).value = Val(Mid(strPar, 1, 1))
    chkDrug(1).value = Val(Mid(strPar, 2, 1))
    chkDrug(2).value = Val(Mid(strPar, 3, 1))
    If Not chkDrug(0).Enabled Then chkDrug(0).Tag = "1" '����û��Ȩ��
    Call SetDrugEnabled 'Ҫ��ִ��
End Sub

Private Sub SetDrugEnabled()
    Dim blnEnabled As Boolean, i As Integer
    
    For i = 0 To lstClass.ListCount - 1
        If lstClass.Selected(i) Then
            If InStr(",5,6,8,", Chr(lstClass.ItemData(i))) > 0 Then
                blnEnabled = True: Exit For
            End If
        End If
    Next
    
    chkDrug(0).Enabled = blnEnabled And chkDrug(0).Tag = ""
    chkDrug(1).Enabled = blnEnabled And chkDrug(0).Tag = ""
    chkDrug(2).Enabled = blnEnabled And chkDrug(0).Tag = ""
End Sub

Private Function Load�������() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str���s As String
    
    On Error GoTo errH
    
    str���s = zlDatabase.GetPara("סԺ�����������", glngSys, pסԺҽ���´�, , Array(lblClass, lstClass))
    
    strSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('7','9') Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        lstClass.AddItem rsTmp!����
        lstClass.ItemData(lstClass.NewIndex) = Asc(rsTmp!����)
        If str���s <> "" Then
            If InStr(str���s, "'" & rsTmp!���� & "'") > 0 Then
                lstClass.Selected(lstClass.NewIndex) = True
            End If
        Else
            lstClass.Selected(lstClass.NewIndex) = True
        End If
        rsTmp.MoveNext
    Next
    If lstClass.ListCount > 0 Then lstClass.ListIndex = 0
    Load������� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    '������������
    If mblnOK Then
        Call zlDatabase.SetPara("סԺ�����������", mstr���s, glngSys, pסԺҽ���´�)
        Call zlDatabase.SetPara("סԺ����ҩƷ�������", mstrҩƷ, glngSys, pסԺҽ���´�)
    End If
End Sub

Private Sub lstClass_ItemCheck(Item As Integer)
    If Visible Then Call SetDrugEnabled
End Sub
