VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmҽ������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ⱥ�����"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frmҽ�������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk���� 
      Caption         =   "�޷ⶥ��(&T)����������Ⱥ��ͳ�ﱨ���ⶥ������"
      Height          =   210
      Index           =   2
      Left            =   825
      TabIndex        =   2
      Top             =   1380
      Width           =   4275
   End
   Begin VB.TextBox txt���� 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "1"
      Top             =   1845
      Width           =   450
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "������(&S)����������Ⱥ��ͳ�ﱨ����������"
      Height          =   210
      Index           =   1
      Left            =   825
      TabIndex        =   1
      Top             =   1065
      Width           =   4275
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "ȫ��ͳ��(&A)����������Ⱥ���б�����Ŀ�������Ը�����"
      Height          =   210
      Index           =   0
      Left            =   825
      TabIndex        =   0
      Top             =   735
      Width           =   4785
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   3870
      MaxLength       =   16
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3210
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh�ֶ� 
      Height          =   1395
      Left            =   825
      TabIndex        =   5
      Top             =   2160
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2461
      _Version        =   393216
      Cols            =   4
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Frame fraTop 
      Height          =   30
      Left            =   -60
      TabIndex        =   12
      Top             =   585
      Width           =   7125
   End
   Begin VB.Frame fraButtom 
      Height          =   30
      Left            =   -30
      TabIndex        =   11
      Top             =   3690
      Width           =   7125
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   315
      TabIndex        =   10
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3255
      TabIndex        =   8
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4500
      TabIndex        =   9
      Top             =   3825
      Width           =   1100
   End
   Begin MSComCtl2.UpDown ud���� 
      Height          =   300
      Left            =   2280
      TabIndex        =   4
      Top             =   1845
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txt����"
      BuddyDispid     =   196610
      OrigLeft        =   2250
      OrigTop         =   1875
      OrigRight       =   2490
      OrigBottom      =   2175
      Max             =   9
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label lblSect 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "�ֶ���Ŀ(&N)"
      Height          =   180
      Left            =   825
      TabIndex        =   13
      Top             =   1905
      Width           =   990
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "������ع������÷ֶΣ��Ա��һ�����÷ֶ�֧��������"
      Height          =   180
      Left            =   840
      TabIndex        =   6
      Top             =   225
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmҽ�������.frx":000C
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmҽ�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlng���� As Long, mlng���� As Long, mlngIndex As Long
Dim mdbl���ֵ As Double
Dim mstrFormat As String   '��ʽ����
Dim mstrλ�� As String
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���


Private Sub chk����_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub chk����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name & 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSects As String
    Dim lngRow As Long
    
    With msh�ֶ�
        For lngRow = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(lngRow, 1)) = "" Then
                MsgBox "��" & lngRow & "������δ���á�", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
            If lngRow < .Rows - 1 Then
                If Val(.TextMatrix(lngRow, 2)) > Val(.TextMatrix(lngRow, 3)) Then
                    MsgBox "��" & lngRow & "������ֵӦ�������ޡ�", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
            End If
            If lngRow > .FixedRows Then
                If Val(.TextMatrix(lngRow, 2)) <> Val(.TextMatrix(lngRow - 1, 3)) + mdbl���ֵ Then
                    MsgBox "��" & lngRow & "����������һ�����޲�������", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
            End If
            If Val(.TextMatrix(lngRow, 2)) <> 0 And Val(.TextMatrix(lngRow - 1, 3)) <> 0 Then
                If Val(.TextMatrix(lngRow, 2)) > 200 Or Val(.TextMatrix(lngRow - 1, 3)) > 200 Then
                    MsgBox "����������޲��ܳ���200�����飡", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
            End If
            .TextMatrix(lngRow, 1) = Join(Split(.TextMatrix(lngRow, 1), ";"))   'ȥ�������ֹ������";"
            strSects = strSects & Val(.TextMatrix(lngRow, 0)) & ";" & Trim(.TextMatrix(lngRow, 1)) & ";" & Val(.TextMatrix(lngRow, 2)) & ";" & Val(.TextMatrix(lngRow, 3)) & ";"
        Next
    End With
    
    On Error GoTo errHandle
    
    gstrSQL = "zl_���������_Update(" & mlng���� & "," & mlng���� & "," & mlngIndex & "," & _
            chk����(0).Value & "," & chk����(1).Value & "," & chk����(2).Value & ",'" & strSects & "')"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitTable()
    With msh�ֶ�
        .TextMatrix(0, 0) = "��"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "����"
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColWidth(0) = 300
        .ColWidth(1) = 1800
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
    End With
End Sub

Private Sub msh�ֶ�_DblClick()
    With msh�ֶ�
        If .COL = 1 Then txtInput.Alignment = 0
        If .COL = 3 Then txtInput.Alignment = 1
        If .COL = 1 Or .COL = 3 And .Row <> .Rows - 1 Then
            txtInput.Move .Left + .CellLeft - 15, .Top + .CellTop - 15, .ColWidth(.COL) - 15, .RowHeight(.Row) - 15
            txtInput.Text = .TextMatrix(.Row, .COL)
            mstrλ�� = .Row & ";" & .COL
            txtInput.Visible = True
            zlControl.TxtSelAll txtInput
            txtInput.SetFocus
        End If
    End With
End Sub

Private Sub msh�ֶ�_KeyPress(KeyAscii As Integer)
    With msh�ֶ�
        Select Case KeyAscii
        Case 13                 'Enter
            If .COL = .Cols - 1 Then
                If .Row = .Rows - 1 Then
                    '�뿪����
                    Me.cmdOK.SetFocus
                    Exit Sub
                End If
                '��һ��
                .Row = .Row + 1
                .COL = .FixedCols
                .TopRow = .Row
            Else
                '��һ��
                .COL = .COL + 1
            End If
        Case 27                     'ESC�˳�
            Call cmdCancel_Click
        Case 32                     '�ո������༭
            Call msh�ֶ�_DblClick
        Case Else                   '����ֱ�ӽ���༭
            Call msh�ֶ�_DblClick
            If .COL = 1 Then
                txtInput.Text = Chr(KeyAscii)
                txtInput.SelStart = Len(txtInput.Text)
            ElseIf .COL = 3 And (KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 64) Then
                '���ּ�����༭
                txtInput.Text = Chr(KeyAscii)
                txtInput.SelStart = Len(txtInput.Text)
            End If
        End Select
    End With
End Sub

Private Sub msh�ֶ�_RowColChange()
    msh�ֶ�.TopRow = msh�ֶ�.Row
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long, lngCol As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call txtInput_Validate(False)
        If txtInput.Visible Then
            Exit Sub
        Else
            msh�ֶ�.SetFocus
            Call msh�ֶ�_KeyPress(13)
        End If
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        lngRow = Split(mstrλ��, ";")(0)
        lngCol = Split(mstrλ��, ";")(1)
        txtInput.Text = msh�ֶ�.TextMatrix(lngRow, lngCol)
        txtInput.Visible = False
        msh�ֶ�.SetFocus
    Else
        lngCol = Split(mstrλ��, ";")(1)
        If lngCol = 3 And (KeyAscii = 45 Or KeyAscii = 47 Or KeyAscii > 65) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtInput_LostFocus()
    txtInput.Visible = False
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
    Dim lngRow As Long, lngCol As Long
    
    With msh�ֶ�
        If txtInput.Visible = False Then Exit Sub
        lngRow = Split(mstrλ��, ";")(0)
        lngCol = Split(mstrλ��, ";")(1)
        If lngCol = 3 And Val(txtInput.Text) = 0 Then
            MsgBox "���޲���Ϊ0��", vbInformation, gstrSysName
            DoEvents
            Cancel = True
            txtInput.Visible = True
            txtInput.SetFocus
            Exit Sub
        End If
        '��д��Ԫ��ֵ
        mblnChange = True
        Select Case lngCol
            Case 1
                .TextMatrix(lngRow, lngCol) = txtInput.Text
            Case 3
                .TextMatrix(lngRow, lngCol) = Format(Val(txtInput.Text), mstrFormat)
                .TextMatrix(lngRow + 1, 2) = Format(Val(.TextMatrix(lngRow, lngCol)) + mdbl���ֵ, mstrFormat)
        End Select
        txtInput.Visible = False
    End With

End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub ud����_Change()
    Dim lngRow As Long, lngCol As Long
    
    mblnChange = True
    With msh�ֶ�
        .Rows = ud����.Value + 1
        For lngRow = .FixedRows + 1 To .Rows - 1
            .TextMatrix(lngRow, 0) = lngRow
            If Trim(.TextMatrix(lngRow - 1, 3)) <> "" Then
                .TextMatrix(lngRow, 2) = Format(Val(.TextMatrix(lngRow - 1, 3)) + mdbl���ֵ, mstrFormat)
            End If
        Next
        .TextMatrix(.Rows - 1, 3) = ""
    End With
End Sub

Public Function ��������(ByVal lng���� As Long, ByVal lng���� As Long, ByVal lngIndex As Long, ByVal STRNAME As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
        
    mblnOK = False
    mlng���� = lng����
    mlng���� = lng����
    mlngIndex = lngIndex
    
    Call InitTable
    frmҽ�������.Caption = STRNAME & "���������"
    mdbl���ֵ = 1
    gstrSQL = "select ����� as ���,����,����,����,nvl(ȫ��ͳ��,0) as ȫ��ͳ�� ,nvl(������,0) as ������ ,nvl(�޷ⶥ��,0) as �޷ⶥ�� " & _
            " from ���������" & _
            " where ����=[1] and ����=[2] and ��ְ=[3]" & _
            " Order by �����"
    mstrFormat = "###;-###; ; "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����, lng����, lngIndex)
    
    If rsTemp.EOF Then
        ud����.Value = 1
        msh�ֶ�.Rows = 2
        msh�ֶ�.TextMatrix(1, 0) = 1
        msh�ֶ�.TextMatrix(1, 1) = STRNAME
    Else
        ud����.Value = rsTemp.RecordCount
        msh�ֶ�.Rows = rsTemp.RecordCount + 1
        
        chk����(0).Value = IIf(rsTemp("ȫ��ͳ��") = 1, 1, 0)
        chk����(1).Value = IIf(rsTemp("������") = 1, 1, 0)
        chk����(2).Value = IIf(rsTemp("�޷ⶥ��") = 1, 1, 0)
        
        lngRow = 1
        Do Until rsTemp.EOF
            msh�ֶ�.TextMatrix(lngRow, 0) = lngRow
            msh�ֶ�.TextMatrix(lngRow, 1) = rsTemp("����")
            msh�ֶ�.TextMatrix(lngRow, 2) = Format(rsTemp("����"), mstrFormat)
            msh�ֶ�.TextMatrix(lngRow, 3) = Format(rsTemp("����"), mstrFormat)
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End If
    
    mblnChange = False
    frmҽ�������.Show vbModal, frmҽ�����
    �������� = mblnOK
End Function
