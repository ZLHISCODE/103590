VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm����֧������ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���֧������"
   ClientHeight    =   5445
   ClientLeft      =   1770
   ClientTop       =   2235
   ClientWidth     =   8250
   Icon            =   "frm����֧������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraButtom 
      Height          =   30
      Left            =   -195
      TabIndex        =   5
      Top             =   4845
      Width           =   9345
   End
   Begin VB.TextBox txtInput 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3660
      MaxLength       =   16
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
      Height          =   4035
      Left            =   330
      TabIndex        =   0
      Top             =   690
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   7117
      _Version        =   393216
      BackColorBkg    =   -2147483643
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   195
      TabIndex        =   3
      Top             =   4995
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5730
      TabIndex        =   1
      Top             =   4995
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6945
      TabIndex        =   2
      Top             =   4995
      Width           =   1100
   End
   Begin VB.Image imgTop 
      Height          =   480
      Left            =   150
      Picture         =   "frm����֧������.frx":000C
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "����ҽ�Ʊ���2002��ҽ������֧���ֶα�������"
      Height          =   180
      Left            =   780
      TabIndex        =   6
      Top             =   300
      Width           =   3780
   End
End
Attribute VB_Name = "frm����֧������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlng���� As Long, mlng���� As Long, mlng��� As Long
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���
Dim mstrλ�� As String

Private Sub cmdHelp_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long, lngCol As Long
    Dim lng����� As Long, lngPreRow As Long
    
    Dim strSects As String
    
    strSects = ""
    With msh����
        For lngRow = .FixedRows To .Rows - 1
            For lngCol = .FixedCols To .Cols - 1
                If Val(.TextMatrix(lngRow, lngCol)) < 0 Or Val(.TextMatrix(lngRow, lngCol)) > 100 Then
                    MsgBox .TextMatrix(lngRow, 0) & .TextMatrix(0, lngCol) & "�ı���δ������ȷ��", vbInformation, gstrSysName
                    .Row = lngRow: .Col = lngCol: .SetFocus: Exit Sub
                End If
            Next
        Next
    End With
    
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    gstrSQL = "ZL_���ձ�������_DELETE(1," & mlng���� & "," & mlng���� & "," & mlng��� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    With msh����
        For lngRow = .FixedRows To .Rows - 1
            For lngCol = .FixedCols To .Cols - 1
                gstrSQL = "ZL_���ձ�������_INSERT(1," & mlng���� & "," & mlng���� & "," & mlng��� & _
                "," & IIf(.TextMatrix(1, lngCol) = "��Ժ", 1, 2) & "," & .ColData(lngCol) & _
                "," & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, lngCol)) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Next
        Next
    End With
    gcnOracle.CommitTrans
    
    mblnOK = True
    mblnChange = False
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub msh����_DblClick()
    With msh����
        If .Col = 0 Then Exit Sub
        txtInput.Alignment = 1
        txtInput.Move .Left + .CellLeft - 15, .Top + .CellTop - 15, .ColWidth(.Col) - 15, .RowHeight(.Row) - 15
        txtInput.Text = .TextMatrix(.Row, .Col)
        mstrλ�� = .Row & ";" & .Col
        txtInput.Visible = True
        zlControl.TxtSelAll txtInput
        txtInput.SetFocus
    End With
End Sub

Private Sub msh����_KeyPress(KeyAscii As Integer)
    With msh����
        Select Case KeyAscii
        Case 13                 'Enter
            If .Col = .Cols - 1 Then
                If .Row = .Rows - 1 Then
                    '�뿪����
                    Me.cmdOK.SetFocus
                    Exit Sub
                End If
                '��һ��
                .Row = .Row + 1
                .Col = .FixedCols
                .TopRow = .Row
            Else
                '��һ��
                .Col = .Col + 1
            End If
        Case 27                     'ESC�˳�
            Call cmdCancel_Click
        Case 32                     '�ո������༭
            Call msh����_DblClick
        Case Else                   '����ֱ�ӽ���༭
            Call msh����_DblClick
            If .Col <> 0 And (KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 64) Then
                '���ּ�����༭
                txtInput.Text = Chr(KeyAscii)
                txtInput.SelStart = Len(txtInput.Text)
            End If
        End Select
    End With
End Sub

Private Sub msh����_RowColChange()
    msh����.TopRow = msh����.Row
    msh����.LeftCol = msh����.Col
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long, lngCol As Long

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call txtInput_Validate(False)
        If txtInput.Visible Then
            Exit Sub
        Else
            msh����.SetFocus
            Call msh����_KeyPress(13)
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        lngRow = Split(mstrλ��, ";")(0)
        lngCol = Split(mstrλ��, ";")(1)
        txtInput.Text = msh����.TextMatrix(lngRow, lngCol)
        txtInput.Visible = False
        msh����.SetFocus
    ElseIf KeyAscii = 45 Or KeyAscii = 47 Or KeyAscii > 65 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtInput_LostFocus()
    txtInput.Visible = False
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
    Dim lngRow As Long, lngCol As Long

    With msh����
        If txtInput.Visible = False Then Exit Sub
        lngRow = Split(mstrλ��, ";")(0)
        lngCol = Split(mstrλ��, ";")(1)
        If Val(txtInput.Text) < 0 Or Val(txtInput.Text) > 100 Then
            MsgBox "����������ڵ���0��С�ڵ���100��", vbInformation, gstrSysName
            DoEvents
            Cancel = True
            txtInput.Visible = True
            zlControl.TxtSelAll txtInput
            txtInput.SetFocus
            Exit Sub
        End If
        '��д��Ԫ��ֵ
        mblnChange = True
        .TextMatrix(lngRow, lngCol) = Format(Val(txtInput.Text), "0.00")
        txtInput.Visible = False
    End With
End Sub

Public Function �༭֧������(ByVal lng���� As Long, ByVal lng���� As Long, ByVal lng��� As Long) As Boolean
'����:��������õĴ��ڽ���ͨѶ�ĳ���
'����ֵ:�༭�ɹ�����True,����ΪFalse
    
    mlng���� = lng����
    mlng���� = lng����
    mlng��� = lng���
    mblnOK = False
    
    Dim lngCount As Integer, lngRow As Long, lngCol As Long
    
    lblNote.Caption = lng��� & "���ҽ������֧���ֶα������ٷֱȣ�����"
    With frm���ձ�������.msh֧������
        msh����.Rows = .Rows
        msh����.Cols = .Cols
        msh����.FixedRows = .FixedRows
        msh����.ColWidth(0) = .ColWidth(0)
        msh����.ColAlignmentFixed(0) = 1
        msh����.Row = 0
        msh����.Col = 0
        msh����.CellAlignment = 4
        For lngCol = msh����.FixedCols To .Cols - 1
            msh����.ColData(lngCol) = .ColData(lngCol)
            msh����.ColWidth(lngCol) = .ColWidth(lngCol)
            msh����.ColAlignmentFixed(lngCol) = 4
            msh����.ColAlignment(lngCol) = 7
        Next
        For lngRow = 0 To .Rows - 1
            msh����.RowData(lngRow) = .RowData(lngRow)
            For lngCol = 0 To .Cols - 1
                msh����.TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
            Next
        Next
        msh����.Row = .FixedRows
        msh����.Col = .FixedCols
        
        msh����.MergeCells = flexMergeFree
        msh����.MergeRow(0) = True
        msh����.MergeCol(0) = True
    End With
    
    mblnChange = False
    frm����֧������.Show vbModal, frm���ձ�������
    �༭֧������ = mblnOK
End Function

