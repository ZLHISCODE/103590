VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCureClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ο�����"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "frmCureClass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   2310
      Left            =   -3315
      TabIndex        =   16
      Tag             =   "1000"
      Top             =   2415
      Visible         =   0   'False
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   4075
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   300
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3030
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1875
      MaxLength       =   8
      TabIndex        =   3
      Tag             =   "����"
      Text            =   "000000"
      Top             =   1350
      Width           =   570
   End
   Begin VB.TextBox txtSymbol 
      Height          =   300
      Left            =   1740
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "����"
      Top             =   2040
      Width           =   1425
   End
   Begin VB.CheckBox chkCodeLen 
      Caption         =   "������ı��볤�ȣ������˵�����ͬ������(&L)"
      Height          =   285
      Left            =   1020
      TabIndex        =   8
      Top             =   2565
      Width           =   4290
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   -15
      TabIndex        =   13
      Top             =   2865
      Width           =   5745
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   9
      Top             =   3045
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   10
      Top             =   3045
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1740
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "����"
      Top             =   1680
      Width           =   3195
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&P"
      Height          =   300
      Left            =   4650
      TabIndex        =   11
      Top             =   690
      Width           =   285
   End
   Begin VB.TextBox txtParent 
      Height          =   300
      Left            =   1740
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "(��)"
      Top             =   690
      Width           =   2895
   End
   Begin VB.TextBox txtUpCode 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "����"
      Text            =   "0000"
      Top             =   1305
      Width           =   975
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   60
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCureClass.frx":000C
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCureClass.frx":05A6
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblHint 
      AutoSize        =   -1  'True
      Caption         =   "(��ʾ����Del����ϼ������ó�������)"
      Height          =   180
      Left            =   1740
      TabIndex        =   18
      Top             =   1035
      Width           =   3150
   End
   Begin VB.Label lblSymbol 
      AutoSize        =   -1  'True
      Caption         =   "����(&S)"
      Height          =   180
      Left            =   1020
      TabIndex        =   6
      Top             =   2100
      Width           =   630
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "��ҩ"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   180
      TabIndex        =   15
      Top             =   555
      Width           =   360
   End
   Begin VB.Label lblNote 
      Caption         =   "�ɸ����ٴ����������Ӧ�ò������ص����Ӧ�òο���ͳһ�������á�"
      Height          =   345
      Left            =   990
      TabIndex        =   14
      Top             =   210
      Width           =   4170
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmCureClass.frx":0B40
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "����(&D)"
      Height          =   180
      Left            =   1020
      TabIndex        =   2
      Top             =   1365
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   1020
      TabIndex        =   4
      Top             =   1725
      Width           =   630
   End
   Begin VB.Label lblParent 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�(&U)"
      Height          =   180
      Left            =   1020
      TabIndex        =   0
      Top             =   750
      Width           =   630
   End
End
Attribute VB_Name = "frmCureClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim intMaxLen As Integer
Dim objNode As Node
Dim mblnCancel As Boolean
Dim mblnChanged As Boolean

Private Sub chkCodeLen_Click()
    If Me.chkCodeLen.Value = 1 Then
        Me.txtCode.MaxLength = intMaxLen - Len(Me.txtUpCode.Text)
    Else
        Me.txtCode.MaxLength = Me.txtCode.Tag
        Me.txtCode.Text = Mid(Me.txtCode.Text, 1, Me.txtCode.MaxLength)
    End If
    Me.txtCode.SetFocus
    If mblnCancel = True Then Exit Sub
    mblnChanged = True
End Sub

Private Sub chkCodeLen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim lngItemID As Long
    
    If Me.txtCode.MaxLength = 0 Then
        MsgBox "�ϼ������Ѿ��ﵽ��󳤶ȣ����������¼���", vbExclamation, gstrSysName
        Me.cmdCancel.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtCode.Text) = "" Then
        MsgBox "�����������", vbExclamation, gstrSysName
        Me.txtCode.SetFocus
        Exit Sub
    End If
    If Me.chkCodeLen.Value = 0 And Len(Trim(Me.txtCode.Text)) <> Me.txtCode.MaxLength Then
        MsgBox "���볤�ȱ���Ϊ" & Me.txtCode.MaxLength & "λ��������ѡ����ĳ���ѡ��", vbExclamation, gstrSysName
        Me.txtCode.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "���Ʊ�������", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "���Ƴ���" & Me.txtName.MaxLength & "�ĳ�������", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    Err = 0: On Error GoTo ErrHand
    If Me.Tag = "����" Then
        lngItemID = zlDatabase.GetNextId("���Ʋο�����")
        gstrSql = "zl_���Ʋο�����_Insert(" & _
            lngItemID & "," & _
            Me.txtParent.Tag & "," & _
            "'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "'," & _
            "'" & Trim(Me.txtName.Text) & "'," & _
            "'" & Trim(Me.txtSymbol.Text) & "'," & _
            Me.lblKind.Tag & "," & _
            Me.chkCodeLen.Value & ")"
    Else
        lngItemID = Me.Tag
        gstrSql = "zl_���Ʋο�����_Update(" & _
            lngItemID & "," & _
            Me.txtParent.Tag & "," & _
            "'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "'," & _
            "'" & Trim(Me.txtName.Text) & "'," & _
            "'" & Trim(Me.txtSymbol.Text) & "'," & _
            Me.chkCodeLen.Value & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    mblnChanged = False
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    With Me.tvwClass
        .Left = Me.txtParent.Left
        .Top = Me.txtParent.Top + Me.txtParent.Height
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    With Me.lblKind
        Select Case Val(.Tag)
        Case 1
            .Caption = "��ҩ"
            Me.Caption = "��ҩ�ο�����༭"
            Me.lblNote.Caption = "��ҩ�ο�ͨ���ɰ���ҩƷҩ�����úͻ�ѧ���ã�����ٴ�������ҩ����Ƚ��з������á�"
        Case 2
            .Caption = "�г�ҩ"
            Me.Caption = "�г�ҩ�ο�����༭"
            Me.lblNote.Caption = "�г�ҩ�ο�ͨ�����Ը�������ɵ���ζ�빦�õȽ��з��࣬Ҳ��ѡ���������෽�����з��ࡣ"
        Case 3
            .Caption = "�в�ҩ"
            Me.Caption = "�в�ҩ�ο�����༭"
            Me.lblNote.Caption = "�в�ҩ�ο�ͨ�����Ը�������ζ�龭���õȽ��з��࣬Ҳ��ѡ������Ȼϵͳ���Եķ��෽����"
        Case 4
            .Caption = "��ҩ�䷽"
            Me.Caption = "��ҩ�䷽�ο�����༭"
            Me.lblNote.Caption = "�䷽��������г�ҩ�ķ��෽�����У������䷽���ٿ��Ը�ϸ�»���Եؽ��С�"
        Case 5
            .Caption = "���Ʋ���"
            Me.Caption = "���Ʋ����ο�����༭"
            Me.lblNote.Caption = "���Ʋ����ο��ɸ����ٴ���ҽ����������Ӧ�ò������ص����ͳһ�������á�"
        End Select
    End With
    
    Err = 0: On Error GoTo ErrHand
    mblnCancel = True

    If Me.Tag = "����" Then
        gstrSql = "select ID,�ϼ�ID,����,����,����" & _
                " From ���Ʋο�����" & _
                " Where ���� = [1] " & _
                " start with �ϼ�ID is null" & _
                " connect by prior ID=�ϼ�ID"
    Else
        gstrSql = "select ID,�ϼ�ID,����,����,����" & _
                " From ���Ʋο�����" & _
                " Where ���� = [1] " & _
                " and  id not in (select id from ���Ʋο����� start with ID = [2] connect by prior id=�ϼ�id ) " & _
                " start with �ϼ�ID is null" & _
                " connect by prior ID=�ϼ�ID"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.lblKind.Tag, Me.Tag)
        
    With rsTemp
        intMaxLen = .Fields("����").DefinedSize
        Me.txtName.MaxLength = .Fields("����").DefinedSize
        Me.txtSymbol.MaxLength = .Fields("����").DefinedSize
        
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    If Me.Tag = "����" Then
        Call zlDefaultCode
    End If
    mblnCancel = False
    mblnChanged = False
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub Form_Load()
    mblnChanged = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChanged = True Then
        If MsgBox("�����Ѿ��ı䣬��ȷ��Ҫ�˳���", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub lblCode_Click()
    Me.txtCode.SetFocus
End Sub

Private Sub lblName_Click()
    Me.txtName.SetFocus
End Sub

Private Sub lblParent_Click()
    Me.txtParent.SetFocus
End Sub

Private Sub lblSymbol_Click()
    Me.txtSymbol.SetFocus
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txtParent.SetFocus
    Call zlDefaultCode
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If cmdSelect Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtCode_Change()
    If mblnCancel = True Then Exit Sub
    mblnChanged = True
End Sub

Private Sub txtcode_GotFocus()
    Me.txtCode.SelStart = 0: Me.txtCode.SelLength = 100
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtName_Change()
    If mblnCancel = True Then Exit Sub
    mblnChanged = True
End Sub

Private Sub txtName_GotFocus()
    Call zlCommFun.OpenIme(True)
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.txtSymbol.Text = zlStr.GetCodeByORCL(Me.txtName.Text, True, 10)
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtParent_Change()
    If mblnCancel = True Then Exit Sub
    mblnChanged = True
End Sub

Private Sub txtParent_GotFocus()
    Me.txtParent.SelStart = 0: Me.txtParent.SelLength = 100
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Me.txtParent.Tag = 0
        Call zlDefaultCode
    End If
    Me.txtParent.SelStart = 0: Me.txtParent.SelLength = 100
End Sub

Private Sub txtParent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtSymbol_Change()
    If mblnCancel = True Then Exit Sub
    mblnChanged = True
End Sub

Private Sub txtSymbol_GotFocus()
    Me.txtSymbol.SelStart = 0: Me.txtSymbol.SelLength = 100
End Sub

Private Sub txtSymbol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtUpCode_Change()
    Me.txtCode.Width = txtUpCode.Width - TextWidth(txtUpCode.Text) - 120
    Me.txtCode.Left = txtUpCode.Left + TextWidth(txtUpCode.Text) + 60
    If mblnCancel = True Then Exit Sub
    mblnChanged = True
End Sub

Private Sub zlDefaultCode()
    '-----------------------------------------------------
    '���ܣ�����ѡ����ϼ�ID(�����Me.txtParent.Tag)���������ñ����ȱʡֵ
    '-----------------------------------------------------
    Err = 0: On Error GoTo ErrHand
    
    Me.chkCodeLen.Value = 0
    Me.chkCodeLen.Enabled = True
   
    If Me.txtParent.Tag = 0 Then
        Me.txtParent.Text = "(��)"
        Me.txtUpCode.Text = ""
        gstrSql = "select max(����) as ���� From ���Ʋο����� Where ����=[1] and �ϼ�ID is null "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.lblKind.Tag)
            
        With rsTemp
            If IIf(IsNull(!����), "", !����) = "" Then
                Me.txtCode.Text = "01"
                Me.txtCode.MaxLength = intMaxLen
                Me.txtCode.Tag = Me.txtCode.MaxLength
                Me.chkCodeLen.Value = 1
                Me.chkCodeLen.Enabled = False
            Else
                Me.txtCode.MaxLength = Len(Trim(!����))
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If !���� = String(Me.txtCode.MaxLength, "9") Then
                    If Me.txtCode.MaxLength >= intMaxLen Then
                        MsgBox "������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������", vbExclamation, gstrSysName
                        Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                        Me.chkCodeLen.Value = 0
                        Me.chkCodeLen.Enabled = False
                    Else
                        MsgBox "�������Ѿ��ﵽ�������ƣ������������볤����������Ҫ", vbExclamation, gstrSysName
                        Me.txtCode.Text = "1" & String(Me.txtCode.MaxLength, "0")
                        Me.txtCode.MaxLength = Me.txtCode.MaxLength + 1
                        Me.txtCode.Tag = Me.txtCode.MaxLength
                        Me.chkCodeLen.Value = 1
                    End If
                Else
                    Me.txtCode.Text = Format(Mid(!����, Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                End If
            End If
        End With
    Else
        With Me.tvwClass
            .Nodes("_" & Me.txtParent.Tag).Selected = True
            Me.txtParent.Text = .SelectedItem.Text
            Me.txtUpCode.Text = Mid(Split(.SelectedItem.Text, "]")(0), 2)
            If .SelectedItem.Children = 0 Then
                Me.txtCode.MaxLength = intMaxLen - Len(Me.txtUpCode.Text)
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If Me.txtCode.MaxLength > 1 Then
                    Me.txtCode.Text = "01"
                Else
                    Me.txtCode.Text = "1"
                End If
                Me.chkCodeLen.Value = 1
                Me.chkCodeLen.Enabled = False
            Else
                gstrSql = "select nvl(max(����),'') as ����  From ���Ʋο����� Where ����=[1] and �ϼ�ID=[2] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.lblKind.Tag, Val(Mid(.SelectedItem.Key, 2)))
                    
                With rsTemp
                    Me.txtCode.MaxLength = Len(!����) - Len(Me.txtUpCode.Text)
                    Me.txtCode.Tag = Me.txtCode.MaxLength
                    If Mid(!����, Len(Me.txtUpCode.Text) + 1) = String(Me.txtCode.MaxLength, "9") Then
                        If Len(Me.txtUpCode.Text) + Me.txtCode.MaxLength >= intMaxLen Then
                            MsgBox "�÷����¼�������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������", vbExclamation, gstrSysName
                            Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                            Me.chkCodeLen.Value = 0
                            Me.chkCodeLen.Enabled = False
                        Else
                            MsgBox "�÷����¼��������Ѿ��ﵽ�������ƣ������������볤����������Ҫ", vbExclamation, gstrSysName
                            Me.txtCode.Text = "1" & String(Me.txtCode.MaxLength, "0")
                            Me.txtCode.MaxLength = Me.txtCode.MaxLength + 1
                            Me.txtCode.Tag = Me.txtCode.MaxLength
                            Me.chkCodeLen.Value = 1
                        End If
                    Else
                        Me.txtCode.Text = Format(Mid(!����, Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                    End If
                End With
            End If
        End With
    End If
    Me.txtParent.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
