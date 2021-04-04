VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiagProof 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����֤������"
   ClientHeight    =   3405
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ListView lvwList 
      Height          =   3375
      Left            =   4770
      TabIndex        =   8
      Top             =   45
      Visible         =   0   'False
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   5953
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
   Begin VB.CheckBox chkStrict 
      Caption         =   "�ϸ���ѭ��ҽ��֤��׼(S)"
      Height          =   255
      Left            =   1035
      TabIndex        =   2
      Top             =   1875
      Value           =   1  'Checked
      Width           =   2745
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1035
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "����"
      Top             =   1395
      Width           =   3510
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3585
      TabIndex        =   4
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2355
      TabIndex        =   3
      Top             =   2880
      Width           =   1100
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   30
      TabIndex        =   6
      Top             =   2700
      Width           =   5745
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   345
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1635
      Top             =   3210
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
            Picture         =   "frmDiagProof.frx":0000
            Key             =   "ITEM"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "֤��(&N)"
      Height          =   180
      Left            =   1035
      TabIndex        =   0
      Top             =   1125
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frmDiagProof.frx":0452
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "��ҽ������Ҫ����֤�����Ӷ��ﵽ׼ȷ�����Ρ����������ҽ��֤��׼ѡ���Ӧ֤��"
      Height          =   345
      Left            =   1035
      TabIndex        =   7
      Top             =   210
      Width           =   3825
   End
End
Attribute VB_Name = "frmDiagProof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem

Public strLefts As String   '�Ѿ����ڵ�ǰ���֤��
Public strRights As String  '�Ѿ����ڵĺ����֤��
Public strProof As String   '�༭������֤��
Dim intCount As Integer

Private Sub chkStrict_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    strProof = ""
    Me.Hide
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim aryItems() As String
    
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "֤���������", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtName.Tag) = "" And Me.chkStrict.Value = 1 Then
        MsgBox "Ҫ��֤��������׼�����Ǻ�", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "֤�򳬹�" & Me.txtName.MaxLength & "�ĳ�������", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    '�ظ��Լ��
    aryItems = Split(Mid(strLefts & strRights, 2), ";")
    For intCount = LBound(aryItems) To UBound(aryItems)
        If Split(aryItems(intCount), ",")(2) = Trim(Me.txtName.Text) Then
            MsgBox "��֤���Ѿ������ڲο���", vbExclamation, gstrSysName
            Me.txtName.SetFocus
            Exit Sub
        End If
    Next
    '���涨��ʽ��֯�༭����Ŀ
    strProof = Me.txtName.Tag & "," & Me.Tag & "," & Trim(Me.txtName.Text)
    Me.Hide
End Sub

Private Sub Form_Load()
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "����", "����", 3000
        .Add , "����", "����", 900
    End With
    Me.lvwList.ColumnHeaders("����").Position = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.lvwList.Visible Then
        Me.lvwList.Visible = False
        Cancel = True
    End If
End Sub

Private Sub lvwList_DblClick()
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwList
        Me.txtName.Tag = Mid(.SelectedItem.Key, 2)
        Me.txtName = .SelectedItem.Text
        Me.txtName.SetFocus
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

Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Exit Sub
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select ID,����,����,����" & _
            " from ��������Ŀ¼" & _
            " where ���='Z'" & _
            "   and (���� like [1] " & _
            "       OR ���� like [2] " & _
            "       OR ���� like [2])" & _
            " And (����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd'))"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Me.txtName.Text) & "%", gstrMatch & Trim(Me.txtName.Text) & "%")
    
    With rsTemp
        If .RecordCount = 0 Then
            If Me.chkStrict.Value = 1 Then
                MsgBox "δ�ҵ�ָ����׼֤�����", vbExclamation, gstrSysName
                Me.txtName.SetFocus
                Exit Sub
            Else
                Me.txtName.Tag = 0
                KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            End If
        End If
        If .RecordCount = 1 Then
            Me.txtName.Tag = !ID
            Me.txtName.Text = IIf(IsNull(!����), "", !����)
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        
        Me.lvwList.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !ID, !����, "ITEM", "ITEM")
            objItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1) = !����
            .MoveNext
        Loop
        With Me.lvwList
            .ListItems(1).Selected = True
            .Left = Me.txtName.Left + 180
            .Width = Me.ScaleWidth - .Left
            .Top = Me.ScaleTop
            .Height = Me.ScaleHeight
            .Visible = True
            .SetFocus
        End With
    End With
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

