VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm����֧���޶� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���֧���޶�"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frm����֧���޶�.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtסԺ���� 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   300
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   12
      Text            =   "1"
      Top             =   1020
      Width           =   330
   End
   Begin VB.Frame fraTop 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   570
      Width           =   7125
   End
   Begin VB.Frame fraButtom 
      Height          =   30
      Left            =   -195
      TabIndex        =   8
      Top             =   4320
      Width           =   7125
   End
   Begin VB.TextBox txt�ⶥ�� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1920
      TabIndex        =   0
      Top             =   675
      Width           =   1005
   End
   Begin VB.TextBox txtInput 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3270
      MaxLength       =   16
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3675
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   195
      TabIndex        =   5
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1995
      TabIndex        =   2
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3210
      TabIndex        =   3
      Top             =   4470
      Width           =   1100
   End
   Begin MSComCtl2.UpDown udסԺ���� 
      Height          =   300
      Left            =   2250
      TabIndex        =   11
      Top             =   1020
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtסԺ����"
      BuddyDispid     =   196609
      OrigLeft        =   2250
      OrigTop         =   1020
      OrigRight       =   2490
      OrigBottom      =   1320
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
      Height          =   2775
      Left            =   930
      TabIndex        =   1
      Top             =   1380
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4895
      _Version        =   393216
      BackColorBkg    =   -2147483643
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image imgTop 
      Height          =   480
      Left            =   150
      Picture         =   "frm����֧���޶�.frx":000C
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "����ҽ�Ʊ���2002��ҽ������֧���޶����"
      Height          =   180
      Left            =   780
      TabIndex        =   9
      Top             =   270
      Width           =   3420
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      Caption         =   "1)��ȷⶥ��             Ԫ"
      Height          =   180
      Left            =   735
      TabIndex        =   7
      Top             =   750
      Width           =   2430
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      Caption         =   "2)�������(        ��סԺ�����)"
      Height          =   180
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   3420
   End
End
Attribute VB_Name = "frm����֧���޶�"
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
    Dim lngRow As Long
    
    Dim strSects As String
    If Val(txt�ⶥ��.Text) <= 0 Then
        MsgBox "�ⶥ�߽��������0��", vbInformation, gstrSysName
        txt�ⶥ��.SetFocus
        Exit Sub
    End If
    If Val(txt�ⶥ��.Text) > 10000000 Then
        MsgBox "�ⶥ�߽��ܴ���1000��", vbInformation, gstrSysName
        txt�ⶥ��.SetFocus
        Exit Sub
    End If
    
    strSects = "A;" & Val(txt�ⶥ��.Text) & ";"
    With msh����
        For lngRow = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngRow, 1)) <= 0 Or Val(.TextMatrix(lngRow, 1)) > 100000 Then
                MsgBox "��" & lngRow & "��סԺ���߽��δ������ȷ��", vbInformation, gstrSysName
                .SetFocus
                Exit Sub
            End If
            
            If lngRow > .FixedRows And Val(.TextMatrix(lngRow, 1)) > Val(.TextMatrix(lngRow - 1, 1)) Then
                MsgBox "��" & lngRow & "��סԺ���߽�����һ�λ���", vbInformation, gstrSysName
                .SetFocus
                Exit Sub
            End If
            strSects = strSects & .TextMatrix(lngRow, 0) & ";" & Val(.TextMatrix(lngRow, 1)) & ";"
        Next
    End With
    
    On Error GoTo ErrHand
    gstrSQL = "zl_����֧���޶�_Update(" & mlng���� & "," & mlng���� & "," & mlng��� & ",'" & strSects & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mblnOK = True
    mblnChange = False
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
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
        txtInput.ZOrder
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
            If .Col = 1 And (KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 64) Then
                '���ּ�����༭
                txtInput.Text = Chr(KeyAscii)
                txtInput.SelStart = Len(txtInput.Text)
            End If
        End Select
    End With
End Sub

Private Sub msh����_RowColChange()
    msh����.TopRow = msh����.Row
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
            Call msh����_KeyPress(vbKeyReturn)
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
        If Val(txtInput.Text) <= 0 Or Val(txtInput.Text) > 100000 Then
            MsgBox "���߽��������0��С��10��", vbInformation, gstrSysName
            DoEvents
            Cancel = True
            txtInput.Visible = True
            txtInput.SetFocus
            Exit Sub
        End If
        '��д��Ԫ��ֵ
        mblnChange = True
        .TextMatrix(lngRow, lngCol) = Format(Val(txtInput.Text), "0")
        txtInput.Visible = False
    End With

End Sub

Private Sub txtסԺ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txt�ⶥ��_Change()
    mblnChange = True
End Sub

Private Sub txt�ⶥ��_GotFocus()
    zlControl.TxtSelAll txt�ⶥ��
End Sub

Private Sub txt�ⶥ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txt�ⶥ��_LostFocus()
    txt�ⶥ��.Text = Format(txt�ⶥ��.Text, "0")
End Sub

Private Sub txt�ⶥ��_Validate(Cancel As Boolean)
    If Val(txt�ⶥ��.Text) <= 0 Then
        Cancel = True
        MsgBox "�ⶥ�߽��������0��", vbInformation, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub udסԺ����_Change()
    Dim lngRow As Long
    
    With msh����
        .Rows = udסԺ����.Value + 1
        lngRow = .Rows - 1
        .TextMatrix(lngRow, 0) = lngRow
        If Trim(.TextMatrix(lngRow, 1)) = "" Then
            .TextMatrix(lngRow, 1) = 0
        End If
    End With
End Sub

Public Function �༭֧���޶�(ByVal lng���� As Long, ByVal lng���� As Long, ByVal lng��� As Long) As Boolean
'����:��������õĴ��ڽ���ͨѶ�ĳ���
'����ֵ:�༭�ɹ�����True,����ΪFalse
    
    mlng���� = lng����
    mlng���� = lng����
    mlng��� = lng���
    mblnOK = False
    
    Dim lngRow As Long
    
    lblNote.Caption = lng��� & "���ҽ������֧���޶����"
    
    With msh����
        .TextMatrix(0, 0) = "סԺ����"
        .TextMatrix(0, 1) = "���"
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignment(0) = 4
        .ColAlignment(1) = 7
        .ColWidth(0) = 1000
        .ColWidth(1) = 1200
    End With
    With frm�������.msh֧���޶�
        txt�ⶥ��.Text = .TextMatrix(.Rows - 1, 1)
        udסԺ����.Value = .Rows - 2
        msh����.Rows = .Rows - 1
        
        For lngRow = 1 To .Rows - 2
            msh����.TextMatrix(lngRow, 0) = lngRow
            msh����.TextMatrix(lngRow, 1) = .TextMatrix(lngRow, 1)
        Next
    End With
    
    mblnChange = False
    frm����֧���޶�.Show vbModal, frm�������
    �༭֧���޶� = mblnOK
End Function
