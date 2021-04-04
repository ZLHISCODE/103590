VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPatiSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "frmPatiSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkOut 
      Caption         =   "��ʾ��Ժ����(&O)"
      Height          =   195
      Left            =   255
      TabIndex        =   5
      Top             =   4875
      Width           =   1650
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3195
      TabIndex        =   4
      Top             =   4800
      Width           =   1150
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2040
      TabIndex        =   3
      Top             =   4800
      Width           =   1150
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgdPati 
      Height          =   4245
      Left            =   75
      TabIndex        =   2
      Top             =   435
      Width           =   4395
      _ExtentX        =   7752
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
   Begin VB.ComboBox cboSect 
      Height          =   300
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   75
      Width           =   3240
   End
   Begin VB.Label lblSect 
      AutoSize        =   -1  'True
      Caption         =   "סԺ����(&D)"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   990
   End
End
Attribute VB_Name = "frmPatiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mobjInput As Object
Private mrsPati As New ADODB.Recordset

Private Sub cboSect_Click()
    Dim strSQL As String, i As Integer
    
    fgdPati.Clear
    
    On Error GoTo errHandle
    If chkOut.Value = 0 Then
        '��ǰ��Ժ����
        strSQL = "Select ����ID,סԺ��,����,��ǰ���� as ��λ,�Ա�,'��' as ��Ժ" & _
                " From ������Ϣ" & _
                " Where ��ǰ����ID = [1] " & _
                " Order by סԺ�� Desc"
    Else
        'ס(��)Ժ����
        strSQL = "Select I.����ID,I.סԺ��,I.����,P.��Ժ���� as ��λ,I.�Ա�," & _
                " Decode(P.��Ժ����,NULL,'��','') as ��Ժ " & _
                " From ������Ϣ I,������ҳ P" & _
                " Where I.����ID=P.����ID And I.סԺ����=P.��ҳID " & _
                " And P.��Ժ����ID = [1] " & _
                " Order by I.סԺ�� Desc"
    End If
    Set mrsPati = zldatabase.OpenSQLRecord(strSQL, Me.Caption, cboSect.ItemData(cboSect.ListIndex))
    
    With mrsPati
        If .RecordCount > 0 Then
            Set fgdPati.Recordset = mrsPati
            fgdPati.ColWidth(0) = 800
            fgdPati.ColWidth(1) = 800
            fgdPati.ColWidth(2) = 850
            fgdPati.ColWidth(3) = 600
            fgdPati.ColWidth(4) = 500
            fgdPati.ColWidth(5) = 500
            fgdPati.ColAlignment(4) = 4
            fgdPati.ColAlignment(5) = 4
        Else
            fgdPati.Rows = 2
            fgdPati.Cols = 2
        End If
    End With
    
    For i = 0 To fgdPati.Cols - 1
        fgdPati.ColAlignmentFixed(i) = 4
    Next
    fgdPati.RowHeight(0) = 320
    fgdPati.Row = 1: fgdPati.TopRow = 1
    fgdPati.Col = 0: fgdPati.ColSel = fgdPati.Cols - 1
    If Visible Then fgdPati.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkOut_Click()
    If cboSect.ListIndex <> -1 Then Call cboSect_Click
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If fgdPati.TextMatrix(1, 0) <> "" Then
        mobjInput.Text = "-" & fgdPati.TextMatrix(fgdPati.Row, 0)
        Unload Me
    End If
End Sub

Private Sub fgdPati_DblClick()
    cmdOK_Click
End Sub

Private Sub fgdPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub Form_Activate()
    fgdPati.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    cboSect.Clear
    With rsTmp
        strSQL = "Select D.ID,D.����,D.���� " & _
            " From ���ű� D,��������˵�� N " & _
            " Where D.ID=N.����ID and N.��������='�ٴ�' and N.������� IN (2,3)" & _
            " And (D.����ʱ�� is NULL or D.����ʱ��=TO_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by D.����"
        .CursorLocation = adUseClient
        Call SQLTest(App.ProductName, Me.Caption, strSQL) 'SQLTest
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, "Form_Load")
        Call SQLTest
        
        Do While Not .EOF
            cboSect.AddItem !���� & "-" & !����
            cboSect.ItemData(cboSect.NewIndex) = !ID
            If !ID = UserInfo.����ID Then cboSect.ListIndex = cboSect.NewIndex
            .MoveNext
        Loop
    End With
    If cboSect.ListCount > 0 And cboSect.ListIndex = -1 Then cboSect.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lblSect_Click()
    cboSect.SetFocus
End Sub

Private Sub fgdPati_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyLeft Then
        If cboSect.ListIndex <> -1 Then
            If cboSect.ListIndex - 1 >= 0 Then
                cboSect.ListIndex = cboSect.ListIndex - 1
                fgdPati.Row = 1: fgdPati.Col = 0: fgdPati.ColSel = fgdPati.Cols - 1: fgdPati.SetFocus
            End If
        End If
    ElseIf KeyCode = vbKeyRight Then
        If cboSect.ListIndex <> -1 Then
            If cboSect.ListIndex + 1 <= cboSect.ListCount - 1 Then
                cboSect.ListIndex = cboSect.ListIndex + 1
                fgdPati.Row = 1: fgdPati.Col = 0: fgdPati.ColSel = fgdPati.Cols - 1: fgdPati.SetFocus
            End If
        End If
    End If
End Sub
