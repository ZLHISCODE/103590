VERSION 5.00
Begin VB.Form frmInAdviceSendCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "frmInAdviceSendCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraDetail 
      Height          =   1920
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   0
      Width           =   4500
      Begin VB.CheckBox chk�Ӱ�Ӽ� 
         Caption         =   "ִ�мӰ�Ӽ�(&V)"
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   1620
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
         Top             =   450
         Width           =   4095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ҫ���͵����(&T):"
         Height          =   180
         Left            =   225
         TabIndex        =   0
         Top             =   225
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3375
      TabIndex        =   4
      Top             =   2010
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2250
      TabIndex        =   3
      Top             =   2010
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
Public mblnOK As Boolean 'OUT:�Ƿ�ȷ��

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
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
    
    gbln�Ӱ�Ӽ� = chk�Ӱ�Ӽ�.Value = 1
    
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
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
    chk�Ӱ�Ӽ�.Value = IIF(gbln�Ӱ�Ӽ�, 1, 0)
    '�������
    Call Load�������
End Sub

Private Function Load�������() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str���s As String
    
    On Error GoTo errH
    
    str���s = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "סԺ�����������", "")
    
    strSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('4','7','9') Order by ����"
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
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "סԺ�����������", mstr���s
    End If
End Sub
