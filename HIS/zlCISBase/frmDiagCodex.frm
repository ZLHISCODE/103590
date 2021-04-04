VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDiagCodex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������������"
   ClientHeight    =   7740
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7560
   Icon            =   "frmDiagCodex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6270
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   900
      Width           =   1200
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   2925
      Left            =   -5565
      TabIndex        =   27
      Top             =   3885
      Visible         =   0   'False
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5159
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "����(&D)"
      Height          =   350
      Index           =   1
      Left            =   5025
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1100
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "����(&U)"
      Height          =   350
      Index           =   0
      Left            =   3930
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1100
   End
   Begin VB.Frame fraDefine 
      Caption         =   "����ϸ����:"
      Height          =   1815
      Left            =   90
      TabIndex        =   25
      Top             =   5895
      Width           =   7395
      Begin VB.ComboBox cboValue 
         Height          =   300
         Left            =   3255
         TabIndex        =   18
         Top             =   915
         Width           =   4035
      End
      Begin VB.ComboBox cboGroup 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   285
         Width           =   1920
      End
      Begin VB.TextBox txtGroup 
         Height          =   300
         Left            =   975
         MaxLength       =   10
         TabIndex        =   12
         Top             =   285
         Width           =   1920
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   90
         TabIndex        =   14
         Top             =   915
         Width           =   1950
      End
      Begin VB.ComboBox cboFormula 
         Height          =   300
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   915
         Width           =   1140
      End
      Begin VB.TextBox txtDegree 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   990
         MaxLength       =   3
         TabIndex        =   20
         Top             =   1350
         Width           =   1050
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "����ϸ��(&A)"
         Height          =   350
         Left            =   6075
         TabIndex        =   21
         Top             =   1350
         Width           =   1200
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ(&I):"
         Height          =   180
         Left            =   90
         TabIndex        =   13
         Top             =   690
         Width           =   720
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "(ͨ�����������ͷ��ڽ��з���)"
         Height          =   180
         Left            =   3015
         TabIndex        =   26
         Top             =   345
         Width           =   2520
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         Caption         =   "������(&N)"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   345
         Width           =   810
      End
      Begin VB.Label lblFormula 
         AutoSize        =   -1  'True
         Caption         =   "����(&F):"
         Height          =   180
         Left            =   2085
         TabIndex        =   15
         Top             =   690
         Width           =   720
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "ֵ(&V):"
         Height          =   180
         Left            =   3255
         TabIndex        =   17
         Top             =   690
         Width           =   540
      End
      Begin VB.Label lblDegree 
         AutoSize        =   -1  'True
         Caption         =   "���ɶ�(&T)"
         Height          =   180
         Left            =   90
         TabIndex        =   19
         Top             =   1410
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "�Ƴ�ϸ��(&R)"
      Height          =   350
      Left            =   6240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1200
   End
   Begin VB.CheckBox chkGroup 
      Caption         =   "��������(&G)"
      Height          =   240
      Left            =   4110
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1155
      Width           =   1710
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6270
      TabIndex        =   24
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6270
      TabIndex        =   23
      Top             =   75
      Width           =   1200
   End
   Begin VB.TextBox txtEnsure 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2805
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "95"
      Top             =   510
      Width           =   795
   End
   Begin VB.TextBox txtUnsure 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2805
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "80"
      Top             =   120
      Width           =   795
   End
   Begin VB.Frame fraCodex 
      Height          =   30
      Left            =   75
      TabIndex        =   22
      Top             =   1005
      Width           =   5460
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -30
      Top             =   7230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagCodex.frx":058A
            Key             =   "ITEM"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdCodex 
      Height          =   4020
      Left            =   90
      TabIndex        =   6
      Top             =   1395
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   7091
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483639
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      MergeCells      =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmDiagCodex.frx":09DC
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblCodex 
      AutoSize        =   -1  'True
      Caption         =   "������������ϸ��(&X):"
      Height          =   180
      Left            =   90
      TabIndex        =   5
      Top             =   1155
      Width           =   1800
   End
   Begin VB.Label lblEnsure 
      AutoSize        =   -1  'True
      Caption         =   "(&2) �����廳�ɶȴﵽ          ʱ����ʾΪ�ٴ���ϡ�"
      Height          =   180
      Left            =   960
      TabIndex        =   2
      Top             =   570
      Width           =   4500
   End
   Begin VB.Label lblUnsure 
      AutoSize        =   -1  'True
      Caption         =   "(&1) �����廳�ɶȴﵽ          ʱ����ʾΪ������ϣ�"
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   180
      Width           =   4500
   End
End
Attribute VB_Name = "frmDiagCodex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngBarSize As Long

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer, lngRow As Integer, lngCol As Integer
Dim blnActive As Boolean

Const conCol������ As Integer = 0
Const conCol��ĿID As Integer = 1
Const conCol��Ŀ�� As Integer = 2
Const conCol��ϵʽ As Integer = 3
Const conCol����ֵ As Integer = 4
Const conCol���ɶ� As Integer = 5

Private Sub cboFormula_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboValue_GotFocus()
    Me.cboValue.SelStart = 0: Me.cboValue.SelLength = 100
End Sub

Private Sub cboValue_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkGroup_Click()
    If Not blnActive Then Exit Sub
    With Me.hgdCodex
        .Redraw = False
        .Clear
        .Rows = 1 + .FixedRows
        .TextMatrix(0, conCol������) = "������"
        .TextMatrix(0, conCol��ĿID) = "��ĿID"
        .TextMatrix(0, conCol��Ŀ��) = "��Ŀ��"
        .TextMatrix(0, conCol��ϵʽ) = "��ϵʽ"
        .TextMatrix(0, conCol����ֵ) = "����ֵ"
        .TextMatrix(0, conCol���ɶ�) = "���ɶ�"
        If Me.chkGroup.Value = 1 Then
            .ColWidth(conCol������) = 1000
            Me.txtGroup.Enabled = True
            Me.txtGroup.BackColor = &H80000005
        Else
            .ColWidth(conCol������) = 0
            Me.txtGroup.Text = ""
            Me.txtGroup.Enabled = False
            Me.txtGroup.BackColor = &H8000000F
        End If
        .ColWidth(conCol����ֵ) = .Width - .ColWidth(conCol������) - .ColWidth(conCol��Ŀ��) - .ColWidth(conCol��ϵʽ) - .ColWidth(conCol���ɶ�) - mlngBarSize
        .Redraw = True
    End With
    MsgBox "����������ʽ�ı䣬�Ѿ����������ȫ������ϸ��", vbExclamation, gstrSysName
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdAppend_Click()
    'ϸ������ȷ�Լ��
    If Me.chkGroup.Value = 1 Then
        If Me.Tag = "��ҽ" Then
            If Trim(Me.txtGroup.Text) = "" Then
                MsgBox "������������˵����������", vbExclamation, gstrSysName
                Me.txtGroup.SetFocus: Exit Sub
            End If
        Else
            If Trim(Me.cboGroup.Text) = "" Then
                MsgBox "��ҽ����˵����֤��������" & vbCrLf & "����޷�ѡ��֤���������ȱ༭�ο�����ȷ��֤��", vbExclamation, gstrSysName
                Me.cboGroup.SetFocus: Exit Sub
            End If
        End If
    End If
    If Trim(Me.txtItem.Tag) <> Trim(Me.txtItem.Text) Or Trim(Me.txtItem.Text) = "" Then
        MsgBox "δָ����ȷ������ϸ����Ŀ��", vbExclamation, gstrSysName
        Me.txtItem.SetFocus: Exit Sub
    End If
    If Trim(Me.cboFormula.Text) = "" Then
        MsgBox "δָ����ȷ������ϸ���ϵʽ��", vbExclamation, gstrSysName
        Me.cboFormula.SetFocus: Exit Sub
    End If
    If Me.cboValue.Enabled Then
        If Trim(Me.cboValue.Text) = "" Then
            MsgBox "δָ��������ϸ������ֵ��", vbExclamation, gstrSysName
            Me.cboValue.SetFocus: Exit Sub
        End If
        strTemp = zlVerifyForm
        If strTemp <> "" Then
            MsgBox strTemp, vbExclamation, gstrSysName
            Me.cboValue.SetFocus: Exit Sub
        End If
    End If
    If Val(Me.txtDegree.Text) = 0 Then
        MsgBox "��Ҫ��ȷ��ϸ���������еĻ��ɶȣ����ܽ�����Ч��������", vbExclamation, gstrSysName
        Me.txtDegree.SetFocus: Exit Sub
    End If
    
    '��ϸ����ӵ�����У��ҵ��뵱ǰϸ��������ͬ�����һ��ϸ���������룬�Ҳ�����������
    Dim intAppendRow As Integer     '��¼����λ�õ���
    With Me.hgdCodex
        .Redraw = False
        If .Rows = 1 + .FixedRows And Val(.TextMatrix(.FixedRows, conCol��ĿID)) = 0 Then
            intAppendRow = .FixedRows
        Else
            intAppendRow = .Rows
            For lngRow = .Rows - 1 To .FixedRows Step -1
                If Me.Tag = "��ҽ" And .TextMatrix(lngRow, conCol������) = Trim(Me.txtGroup.Text) _
                   Or Me.Tag <> "��ҽ" And .TextMatrix(lngRow, conCol������) = Trim(Me.cboGroup.Text) Then
                    intAppendRow = lngRow + 1: Exit For
                End If
            Next
            .Rows = .Rows + 1
            For lngRow = .Rows - 2 To intAppendRow Step -1
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow + 1, lngCol) = .TextMatrix(lngRow, lngCol)
                Next
            Next
        End If
        .TextMatrix(intAppendRow, conCol������) = IIf(Me.Tag = "��ҽ", Trim(Me.txtGroup.Text), Trim(Me.cboGroup.Text))
        .TextMatrix(intAppendRow, conCol��ĿID) = Val(Me.lblItem.Tag)
        .TextMatrix(intAppendRow, conCol��Ŀ��) = Trim(Me.txtItem.Text)
        .TextMatrix(intAppendRow, conCol��ϵʽ) = Trim(Me.cboFormula.Text)
        .TextMatrix(intAppendRow, conCol����ֵ) = Trim(Me.cboValue.Text)
        .TextMatrix(intAppendRow, conCol���ɶ�) = Val(Me.txtDegree.Text)
        .Row = intAppendRow
        .Col = conCol��Ŀ��
        .Redraw = True
    End With
    
    '���ϸ����ؼ����ݣ��Ա㶨���µ�ϸ��
    Me.lblItem.Tag = ""
    Me.txtItem.Text = ""
    Me.txtItem.Tag = ""
    Me.lblValue.Tag = ""
    Me.cboValue.Text = ""
    Me.lblFormula.Tag = ""
    Me.cboFormula.Clear
    Me.hgdCodex.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdMove_Click(Index As Integer)
    With Me.hgdCodex
        If Index = 0 Then
            If .Row = .FixedRows Then Exit Sub
            If .TextMatrix(.Row, conCol������) <> .TextMatrix(.Row - 1, conCol������) Then
                MsgBox "����ϸ�����ڲ�ͬ����֮���ƶ���", vbExclamation, gstrSysName: Exit Sub
            End If
        Else
            If .Row = .Rows - 1 Then Exit Sub
            If .TextMatrix(.Row, conCol������) <> .TextMatrix(.Row + 1, conCol������) Then
                MsgBox "����ϸ�����ڲ�ͬ����֮���ƶ���", vbExclamation, gstrSysName: Exit Sub
            End If
        End If
        
        .Redraw = False
        strTemp = ""
        For lngCol = 0 To .Cols - 1
            strTemp = strTemp & "|" & .TextMatrix(.Row, lngCol)
        Next
        For lngCol = 0 To .Cols - 1
            If Index = 0 Then
                .TextMatrix(.Row, lngCol) = .TextMatrix(.Row - 1, lngCol)
                .TextMatrix(.Row - 1, lngCol) = Split(Mid(strTemp, 2), "|")(lngCol)
            Else
                .TextMatrix(.Row, lngCol) = .TextMatrix(.Row + 1, lngCol)
                .TextMatrix(.Row + 1, lngCol) = Split(Mid(strTemp, 2), "|")(lngCol)
            End If
        Next
        If Index = 0 Then
            .Row = .Row - 1
        Else
            .Row = .Row + 1
        End If
        .Redraw = True
    End With
    Call hgdCodex_RowColChange
End Sub

Private Sub cmdOK_Click()
    Dim intGrpNo As Integer, intItmNo As Integer
    If Val(Me.txtUnsure.Text) = 0 And Val(Me.txtEnsure.Text) = 0 Then
        MsgBox "Ϊ��д������Ϻ��ٴ����Ҫ��Ļ��ɶȣ�", vbExclamation, gstrSysName
        Me.txtUnsure.SetFocus: Exit Sub
    End If
    If Val(Me.txtEnsure.Text) <> 0 And Val(Me.txtEnsure.Text) < Val(Me.txtUnsure.Text) Then
        MsgBox "������ϻ��ɶȲ�Ӧ�����ٴ����Ҫ��Ļ��ɶȣ�", vbExclamation, gstrSysName
        Me.txtUnsure.SetFocus: Exit Sub
    End If
    
    With Me.hgdCodex
        strTemp = "δ����"
        intGrpNo = -1: intItmNo = 0: gstrSql = ""
        For lngRow = .FixedRows To .Rows - 1
            If Trim(Me.hgdCodex.TextMatrix(Me.hgdCodex.FixedRows, conCol��ĿID)) <> "" Then
                If strTemp <> Trim(.TextMatrix(lngRow, conCol������)) Then
                    intGrpNo = intGrpNo + 1: intItmNo = 0
                End If
                intItmNo = intItmNo + 1
                gstrSql = gstrSql & "|" & _
                    intGrpNo & "^" & Trim(.TextMatrix(lngRow, conCol������)) & "^" & _
                    intItmNo & "^" & Trim(.TextMatrix(lngRow, conCol��ĿID)) & "^" & _
                    Trim(.TextMatrix(lngRow, conCol��ϵʽ)) & "^" & _
                    Trim(.TextMatrix(lngRow, conCol����ֵ)) & "^" & _
                    Trim(.TextMatrix(lngRow, conCol���ɶ�))
                strTemp = Trim(.TextMatrix(lngRow, conCol������))
            End If
        Next
    End With
    If gstrSql = "" Then
        If MsgBox("δ�����κ�����ϸ��,���ȷ������ɾ��������ϸ��" & vbCrLf & "������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Me.hgdCodex.SetFocus: Exit Sub
        Else
            gstrSql = "zl_������Ϲ���_Update(" & Me.hgdCodex.Tag & ",0,0,'')"
        End If
    Else
        gstrSql = "zl_������Ϲ���_Update(" & _
                Me.hgdCodex.Tag & "," & _
                Val(Trim(Me.txtUnsure.Text)) & "," & _
                Val(Trim(Me.txtEnsure.Text)) & "," & _
                "'" & Mid(gstrSql, 2) & "')"
    End If
    Err = 0: On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRemove_Click()
    If Val(Me.hgdCodex.TextMatrix(Me.hgdCodex.Row, conCol��ĿID)) = 0 Then Exit Sub
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select id, ����, ������, Ӣ����, ����, ����, С��, ��λ, ��ʾ��,��ֵ��" & _
            " from ����������Ŀ I" & _
            " where id=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdCodex.TextMatrix(Me.hgdCodex.Row, conCol��ĿID)))
    
    With rsTemp
        Me.lblItem.Tag = !ID
        Me.txtItem.Text = !������
        Me.txtItem.Tag = !������
        Me.lblFormula.Tag = IIf(IsNull(!����), 0, !����)
        Me.lblValue.Tag = IIf(IsNull(!��ֵ��), "", !��ֵ��)
        Me.cboValue.Tag = IIf(IsNull(!��λ), "", !��λ)
        Call zlAdjustForm
    End With
    
    Err = 0: On Error GoTo 0
    With Me.hgdCodex
        Me.cboFormula.Text = .TextMatrix(.Row, conCol��ϵʽ)
        Me.cboValue.Text = .TextMatrix(.Row, conCol����ֵ)
        Me.txtDegree.Text = .TextMatrix(.Row, conCol���ɶ�)
        For lngRow = .Row To .Rows - 2
            For lngCol = 0 To .Cols - 1
                .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow + 1, lngCol)
            Next
        Next
        If .Rows = 1 + .FixedRows Then
            For lngCol = 0 To .Cols - 1
                .TextMatrix(.FixedRows, lngCol) = ""
            Next
        Else
            .Rows = .Rows - 1
        End If
    End With
    Call hgdCodex_RowColChange
    Me.txtItem.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If blnActive Then Exit Sub
    
    Err = 0: On Error GoTo ErrHand
    
    '����������д
    gstrSql = "select ID,����,����,�ٴ�" & _
            " from �������Ŀ¼" & _
            " where ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdCodex.Tag))
    
    With rsTemp
        Me.Caption = !���� & "����������"
        Me.txtUnsure.Text = IIf(IsNull(!����), 0, !����)
        Me.txtEnsure.Text = IIf(IsNull(!�ٴ�), 0, !�ٴ�)
    End With
        
    '����ϸ����д
    Me.hgdCodex.Redraw = False
    intCount = 0
    
    gstrSql = "select R.������,R.��ĿID,I.������ as ��Ŀ��,R.��ϵʽ,R.����ֵ,R.���ɶ�" & _
            " from ������Ϲ��� R,����������Ŀ I" & _
            " where R.��ĿID=I.ID and R.���ID=[1] " & _
            " order by R.�����,R.������"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdCodex.Tag))
    
    With rsTemp
        If .EOF Then
            Me.hgdCodex.Rows = 1 + Me.hgdCodex.FixedRows
        Else
            Me.hgdCodex.Rows = .RecordCount + Me.hgdCodex.FixedRows
        End If
        Do While Not .EOF
            If Trim(IIf(IsNull(!������), "", !������)) <> "" Then
                intCount = 1
            End If
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol������) = IIf(IsNull(!������), "", !������)
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol��ĿID) = IIf(IsNull(!��ĿID), 0, !��ĿID)
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol��Ŀ��) = IIf(IsNull(!��Ŀ��), "", !��Ŀ��)
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol��ϵʽ) = IIf(IsNull(!��ϵʽ), "", !��ϵʽ)
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol����ֵ) = IIf(IsNull(!����ֵ), "", !����ֵ)
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol���ɶ�) = IIf(IsNull(!���ɶ�), 0, !���ɶ�)
            .MoveNext
        Loop
    End With
    
    If Me.Tag <> "��ҽ" Then
        gstrSql = "select distinct ֤������" & _
                " from ������ϲο�" & _
                " where ���id=[1] " & _
                "       and ֤������ is not null"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdCodex.Tag))
        
        Me.cboGroup.Clear
        Do While Not rsTemp.EOF
            Me.cboGroup.AddItem rsTemp!֤������
            rsTemp.MoveNext
        Loop
    End If
    
    With Me.hgdCodex
        If Me.Tag = "��ҽ" Then
            Me.txtGroup.Visible = True
            Me.cboGroup.Visible = False
            Me.lblNote.Caption = "(ͨ�����������ͷ��ڽ��з���)"
            Me.chkGroup.Caption = "��������(&G)"
            Me.chkGroup.Enabled = True
            If intCount = 1 Then
                Me.chkGroup.Value = 1
                .ColWidth(conCol������) = 1000
                Me.txtGroup.Enabled = True
                Me.txtGroup.BackColor = &H80000005
                Me.lblNote.Visible = True
            Else
                Me.chkGroup.Value = 0
                .ColWidth(conCol������) = 0
                Me.txtGroup.Enabled = False
                Me.txtGroup.BackColor = &H8000000F
                Me.lblNote.Visible = False
            End If
        Else
            Me.txtGroup.Visible = False
            Me.cboGroup.Visible = True
            Me.lblNote.Caption = "(Ҫ�󰴲ο����Ѿ������ı�֤������з���)"
            Me.lblNote.Visible = True
            Me.chkGroup.Caption = "��֤����(&G)"
            Me.chkGroup.Value = 1
            Me.chkGroup.Enabled = False
            .ColWidth(conCol������) = 1000
        End If
        .ColWidth(conCol����ֵ) = .Width - .ColWidth(conCol������) - .ColWidth(conCol��Ŀ��) - .ColWidth(conCol��ϵʽ) - .ColWidth(conCol���ɶ�) - mlngBarSize
        .Row = .FixedRows
        .Col = conCol��Ŀ��
    End With
    
    Me.hgdCodex.Redraw = True
    blnActive = True
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwList.Visible Then
        Me.lvwList.Visible = False
        Me.txtItem.SetFocus
    Else
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    blnActive = False
    With Me.hgdCodex
        .Redraw = False
        .ColAlignment(conCol������) = 0
        .ColAlignment(conCol��ĿID) = 1
        .ColAlignment(conCol��Ŀ��) = 1
        .ColAlignment(conCol��ϵʽ) = 1
        .ColAlignment(conCol����ֵ) = 1
        .ColAlignment(conCol���ɶ�) = 6
        
        .TextMatrix(0, conCol������) = "������"
        .TextMatrix(0, conCol��ĿID) = "��ĿID"
        .TextMatrix(0, conCol��Ŀ��) = "��Ŀ��"
        .TextMatrix(0, conCol��ϵʽ) = "��ϵʽ"
        .TextMatrix(0, conCol����ֵ) = "����ֵ"
        .TextMatrix(0, conCol���ɶ�) = "���ɶ�"
        .MergeCol(0) = True
        
        .ColWidth(conCol������) = 0
        .ColWidth(conCol��ĿID) = 0
        .ColWidth(conCol��Ŀ��) = 1600
        .ColWidth(conCol��ϵʽ) = 900
        .ColWidth(conCol���ɶ�) = 650
        .ColWidth(conCol����ֵ) = .Width - .ColWidth(conCol������) - .ColWidth(conCol��Ŀ��) - .ColWidth(conCol��ϵʽ) - .ColWidth(conCol���ɶ�) - mlngBarSize
        .Redraw = True
    End With
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "������", "������", 1800
        .Add , "����", "����", 1000
        .Add , "����", "����", 600
        .Add , "��ֵ��", "��ֵ��", 4000
    End With
    Me.lvwList.ColumnHeaders("����").Position = 1
End Sub

Private Sub hgdCodex_GotFocus()
    Call hgdCodex_RowColChange
End Sub

Private Sub hgdCodex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub hgdCodex_RowColChange()
    Dim lngCurRow As Long
    With Me.hgdCodex
        If Me.Tag = "��ҽ" Then
            Me.txtGroup.Text = Left(.TextMatrix(.Row, conCol������), 10)
        Else
            For intCount = 0 To Me.cboGroup.ListCount - 1
                If Me.cboGroup.List(intCount) = .TextMatrix(.Row, conCol������) Then
                    Me.cboGroup.ListIndex = intCount
                End If
            Next
        End If
        .Redraw = False
        lngCurRow = .Row
        For lngRow = .FixedRows To .Rows - 1
            .Row = lngRow
            For lngCol = .FixedCols To .Cols - 1
                .Col = lngCol
                If lngRow = lngCurRow Then
                    .CellBackColor = &H80000001
                    .CellForeColor = &H80000005
                Else
                    .CellBackColor = .BackColor
                    .CellForeColor = .ForeColor
                End If
            Next
        Next
        .Row = lngCurRow
        .Col = conCol��Ŀ��
        .Redraw = True
    End With
End Sub

Private Sub lvwList_DblClick()
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwList
        Me.lblItem.Tag = Mid(.SelectedItem.Key, 2)
        Me.txtItem.Text = Split(.SelectedItem.Tag, ",")(0)
        Me.txtItem.Tag = Split(.SelectedItem.Tag, ",")(0)
        Me.lblFormula.Tag = Split(.SelectedItem.Tag, ",")(1)
        Me.lblValue.Tag = .SelectedItem.SubItems(Me.lvwList.ColumnHeaders("��ֵ��").Index - 1)
        Me.cboValue.Tag = Split(.SelectedItem.Tag, ",")(2)
        Call zlAdjustForm
        Me.txtItem.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        Call lvwList_DblClick
    End Select
End Sub

Private Sub lvwList_LostFocus()
    Me.lvwList.Visible = False
End Sub

Private Sub txtDegree_GotFocus()
    Me.txtDegree.SelStart = 0: Me.txtDegree.SelLength = 100
End Sub

Private Sub txtDegree_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtEnsure_GotFocus()
    Me.txtEnsure.SelStart = 0: Me.txtEnsure.SelLength = 100
End Sub

Private Sub txtEnsure_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtGroup_GotFocus()
    Call zlCommFun.OpenIme(True)
    Me.txtGroup.SelStart = 0: Me.txtGroup.SelLength = 100
End Sub

Private Sub txtGroup_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtGroup_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtItem_GotFocus()
    Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$^&*()+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Err = 0: On Error GoTo ErrHand
    
    '������Ŀ����������Ϊ������Ŀ�����Ա����䡢ְҵ����
    gstrSql = "select I.id, I.����, I.������, I.Ӣ����, I.����, I.����, I.С��, I.��λ, I.��ʾ��,I.��ֵ��" & _
            " from ����������Ŀ I,������������ C" & _
            " where I.����ID=C.ID" & _
            "       And (C.����=1 And C.���� Not In ('01', '03', '04', '05') And I.������ <> '����' Or C.����<>1)" & _
            "       And (I.���� like [1] " & _
            "           or I.������ like [2] " & _
            "           or Upper(I.Ӣ����) like [2])"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Me.txtItem.Text) & "%", gstrMatch & Trim(Me.txtItem.Text) & "%")
    
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "δ�ҵ�ָ������������Ŀ", vbExclamation, gstrSysName
            Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
            Me.txtItem.SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.lblItem.Tag = !ID
            Me.txtItem.Text = !������
            Me.txtItem.Tag = !������
            Me.lblFormula.Tag = IIf(IsNull(!����), 0, !����)
            Me.lblValue.Tag = IIf(IsNull(!��ֵ��), "", !��ֵ��)
            Me.cboValue.Tag = IIf(IsNull(!��λ), "", !��λ)
            Call zlAdjustForm
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            Exit Sub
        End If
        
        Me.lvwList.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !ID, !������ & IIf(IsNull(!Ӣ����), "", "(" & !Ӣ���� & ")"), "ITEM", "ITEM")
            objItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1) = !����
            Select Case IIf(IsNull(!����), 0, !����)
            Case 0
                objItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1) = "��ֵ"
            Case 1
                objItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1) = "����"
            Case 2
                objItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1) = "����"
            Case 3
                objItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1) = "�߼�"
            End Select
            objItem.SubItems(Me.lvwList.ColumnHeaders("��ֵ��").Index - 1) = IIf(IsNull(!��ֵ��), "", !��ֵ��)
            objItem.Tag = !������ & "," & IIf(IsNull(!����), 0, !����) & "," & IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
        With Me.lvwList
            .ListItems(1).Selected = True
            .Left = Me.fraDefine.Left + Me.txtItem.Left
            .Top = Me.fraDefine.Top + Me.txtItem.Top - .Height
            .Visible = True
            .SetFocus
        End With
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtUnsure_GotFocus()
    Me.txtUnsure.SelStart = 0: Me.txtUnsure.SelLength = 100
End Sub

Private Sub txtUnsure_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub zlAdjustForm()
    '-------------------------------------------------
    '�����������ʽ�Ŀ�ѡ��Χ
    '��Σ� ������Me.lblFormula.Tag�е���ֵ���ͣ�Me.lblValue.Tag�е���ֵ��
    '-------------------------------------------------
    Dim aryValue() As String
    Me.cboValue.Clear
    Me.cboValue.Enabled = False
    Me.cboFormula.Clear
    Select Case Val(Me.lblFormula.Tag)
    Case 0  '��ֵ
        If Me.cboValue.Tag = "" Then
            Me.lblValue.Caption = "ֵ(&V):(��ֵ��)"
        Else
            Me.lblValue.Caption = "ֵ(&V):(��ֵ�� ��λ:" & Me.cboValue.Tag & ")"
        End If
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "С��"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Enabled = True
    Case 1  '����
        Me.lblValue.Caption = "ֵ(&V):(������)"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Enabled = True
    Case 2  '����
        Me.lblValue.Caption = "ֵ(&V):(������)"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Enabled = True
    Case 3  '�߼�
        Me.lblValue.Caption = "ֵ(&V):(�߼���)"
        Me.cboFormula.AddItem "��"
        Me.cboFormula.AddItem "��"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Text = ""
        Me.cboValue.Enabled = False
    Case Else
    End Select
    
    aryValue = Split(Me.lblValue.Tag, ";")
    For intCount = LBound(aryValue) To UBound(aryValue)
        Me.cboValue.AddItem aryValue(intCount)
    Next
End Sub

Private Function zlVerifyForm() As String
    '-------------------------------------------------
    '�ж��������ʽ��ֵ�������ȷ��
    '��Σ�������Me.lblFormula.Tag�е���ֵ����
    '       Me.lblValue.Tag�е���ֵ��
    '       Me.lblFormula.text�еĹ�ϵʽ
    '       Me.lblValue.text�е�����
    '���Σ���ȷ����""�����򷵻ش�����Ϣ
    '-------------------------------------------------
    Dim aryValue() As String
    zlVerifyForm = ""
    On Error GoTo ErrHandle
    Select Case Val(Me.lblFormula.Tag)
    Case 0  '��ֵ
        Select Case Me.cboFormula.Text
        Case "����", "������", "����", "С��", "����", "����"
            Me.cboValue.Text = Val(Me.cboValue.Text)
        Case "����"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) <> 1 Then
                zlVerifyForm = "����ֵδ�������ڡ�Ҫ�����ֵ1,ֵ2����ʽ��֯��д��": Exit Function
            End If
            Me.cboValue.Text = Val(aryValue(0)) & "," & Val(aryValue(1))
        Case "����", "������"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) < 1 Then
                zlVerifyForm = "�����Ϊ������ֵ��û��Ҫ���á����ڡ��򡰲����ڡ��Ĺ�ϵʽ��": Exit Function
            End If
            Me.cboValue.Text = ""
            For intCount = LBound(aryValue) To UBound(aryValue)
                Me.cboValue.Text = Me.cboValue.Text & "," & Val(aryValue(intCount))
            Next
            Me.cboValue.Text = Mid(Me.cboValue.Text, 2)
        End Select
    Case 1  '����
        Select Case Me.cboFormula.Text
        Case "����", "������", "����", "������"
        Case "����", "������"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) < 1 Then
                zlVerifyForm = "�����Ϊ������ֵ��û��Ҫ���á����ڡ��򡰲����ڡ��Ĺ�ϵʽ��": Exit Function
            End If
        End Select
    Case 2  '����
        Select Case Me.cboFormula.Text
        Case "����", "������", "����", "����", "������", "������"
            gstrSql = "select to_char(to_date('" & Trim(Me.cboValue.Text) & "','YYYY-MM-DD'),'YYYY-MM-DD') from dual"
            With rsTemp
'                If .State = adStateOpen Then .Close
                Err = 0: On Error Resume Next
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "zlVerifyForm")
                If Err <> 0 Then zlVerifyForm = "��������ֵ���������ڸ�ʽ�涨(YYYY-MM-DD)��": Exit Function
                Err = 0: On Error GoTo 0
                Me.cboValue.Text = .Fields(0).Value
            End With
        Case "����"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) <> 1 Then
                zlVerifyForm = "����ֵδ�������ڡ�Ҫ�����ֵ1,ֵ2����ʽ��֯��д��": Exit Function
            End If
            Me.cboValue.Text = ""
            For intCount = LBound(aryValue) To UBound(aryValue)
                gstrSql = "select to_char(to_date('" & Trim(aryValue(intCount)) & "','YYYY-MM-DD'),'YYYY-MM-DD') from dual"
                With rsTemp
'                    If .State = adStateOpen Then .Close
                    Err = 0: On Error Resume Next
                    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "zlVerifyForm")
                    If Err <> 0 Then zlVerifyForm = "��������ֵ�е�" & intCount + 1 & "��������ڸ�ʽ�涨(YYYY-MM-DD)��": Exit Function
                    Err = 0: On Error GoTo 0
                    aryValue(intCount) = .Fields(0).Value
                End With
                Me.cboValue.Text = Me.cboValue.Text & "," & aryValue(intCount)
            Next
            Me.cboValue.Text = Mid(Me.cboValue.Text, 2)
        End Select
    Case 3  '�߼�
    Case Else
    End Select
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

