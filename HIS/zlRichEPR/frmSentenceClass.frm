VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSentenceClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ʾ�ʾ������"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmSentenceClass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ListBox lstScope 
      Columns         =   2
      Height          =   915
      IntegralHeight  =   0   'False
      ItemData        =   "frmSentenceClass.frx":000C
      Left            =   1575
      List            =   "frmSentenceClass.frx":0013
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   2445
      Width           =   4200
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   2940
      Left            =   6075
      TabIndex        =   16
      Tag             =   "1000"
      Top             =   495
      Visible         =   0   'False
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   5186
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
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   19
      Top             =   3435
      Width           =   6345
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   210
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3525
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1710
      MaxLength       =   8
      TabIndex        =   3
      Tag             =   "����"
      Text            =   "000000"
      Top             =   1020
      Width           =   570
   End
   Begin VB.TextBox txtExplain 
      Height          =   645
      Left            =   1575
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   7
      Tag             =   "����"
      Top             =   1710
      Width           =   4185
   End
   Begin VB.CheckBox chkCodeLen 
      Caption         =   "������ĳ��ȣ�������ͬ������(&L)"
      Height          =   285
      Left            =   2730
      TabIndex        =   10
      Top             =   983
      Width           =   4290
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -45
      TabIndex        =   15
      Top             =   795
      Width           =   6345
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1575
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "����"
      Top             =   1350
      Width           =   4185
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&P"
      Height          =   300
      Left            =   5475
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   150
      Width           =   285
   End
   Begin VB.TextBox txtParent 
      Height          =   300
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "(��)"
      Top             =   150
      Width           =   3900
   End
   Begin VB.TextBox txtUpCode 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1575
      MaxLength       =   8
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "����"
      Text            =   "0000"
      Top             =   975
      Width           =   975
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   105
      Top             =   945
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
            Picture         =   "frmSentenceClass.frx":0021
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceClass.frx":05BB
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   11
      Top             =   3540
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4695
      TabIndex        =   12
      Top             =   3540
      Width           =   1100
   End
   Begin VB.Label lblScope 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Χ(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   855
      TabIndex        =   8
      Top             =   2445
      Width           =   630
   End
   Begin VB.Label lblHint 
      AutoSize        =   -1  'True
      Caption         =   "(��ʾ����Del����ϼ������ó�������)"
      Height          =   180
      Left            =   1575
      TabIndex        =   18
      Top             =   495
      Width           =   3150
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      Caption         =   "˵��(&E)"
      Height          =   180
      Left            =   855
      TabIndex        =   6
      Top             =   1770
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmSentenceClass.frx":0B55
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "����(&D)"
      Height          =   180
      Left            =   855
      TabIndex        =   2
      Top             =   1035
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   855
      TabIndex        =   4
      Top             =   1395
      Width           =   630
   End
   Begin VB.Label lblParent 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�(&U)"
      Height          =   180
      Left            =   855
      TabIndex        =   0
      Top             =   210
      Width           =   630
   End
End
Attribute VB_Name = "frmSentenceClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnAdd As Boolean      '�Ƿ�����
Private mblnOK As Boolean       '�Ƿ�ȷ��
Private mintMaxLen As Integer   '������볤��

Dim lngCount As Long

'-----------------------------------------------------
'����Ϊ�ⲿ��������
'-----------------------------------------------------
Public Function ShowMe(ByVal frmParent As Form, ByVal blnAdd As Boolean, ByVal objEdit As MSComctlLib.Node) As Boolean
    '���ܣ���ʾ���༭����
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       objEdit-��ǰ�༭�Ľڵ�
    Dim rsTemp As New ADODB.Recordset
    Dim objNode As MSComctlLib.Node
    mblnAdd = blnAdd
    
    With Me.lstScope
        .Clear
        .AddItem "1-���ﲡ��": .Selected(0) = True
        .AddItem "2-סԺ����": .Selected(1) = True
        .AddItem "3-�����¼": .Selected(2) = False
        .AddItem "4-������": .Selected(3) = True
        .AddItem "5-����֤���뱨��": .Selected(4) = True
        .AddItem "6-֪���ļ�": .Selected(5) = True
        .AddItem "7-���Ʊ���": .Selected(6) = True
        .AddItem "8-��������": .Selected(7) = True
        .ListIndex = 0
    End With
    
    
    'װ���ѡ��������
    Err = 0: On Error GoTo ErrHand
    If mblnAdd Then
        gstrSQL = "Select ID, �ϼ�id, ����, ����, ˵��" & vbNewLine & _
                "From �����ʾ����" & vbNewLine & _
                "Start With �ϼ�id Is Null" & vbNewLine & _
                "Connect By Prior ID = �ϼ�id" & vbNewLine & _
                "Order By Level, ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Else
        gstrSQL = "Select ID, �ϼ�id, ����, ����, ˵��" & vbNewLine & _
                "From �����ʾ����" & vbNewLine & _
                "Where ID Not In (Select ID From �����ʾ���� Start With �ϼ�id = [1] Connect By Prior ID = �ϼ�id)" & vbNewLine & _
                "Start With �ϼ�id Is Null" & vbNewLine & _
                "Connect By Prior ID = �ϼ�id" & vbNewLine & _
                "Order By Level, ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Mid(objEdit.Key, 2)))
    End If
    With rsTemp
        mintMaxLen = .Fields("����").DefinedSize
        Me.txtName.MaxLength = .Fields("����").DefinedSize
        Me.txtExplain.MaxLength = .Fields("˵��").DefinedSize
        
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, !���� & "-" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, !���� & "-" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    
    If mblnAdd Then
        If objEdit Is Nothing Then
            Me.txtParent.Tag = 0
        Else
            Me.txtParent.Tag = Mid(objEdit.Key, 2)
        End If
        Me.Tag = 0
        Call zlDefaultCode
    Else
        If objEdit.Parent Is Nothing Then
            Me.txtParent.Tag = 0
            Me.txtParent.Text = "(��)"
            Me.txtUpCode.Text = ""
            Me.txtCode.Text = Split(objEdit.Text, "-")(0)
            Me.txtCode.MaxLength = Len(Me.txtCode.Text)
            Me.txtCode.Tag = Me.txtCode.MaxLength
        Else
            Me.txtParent.Tag = Mid(objEdit.Parent.Key, 2)
            Me.txtParent.Text = objEdit.Parent.Text
            Me.txtUpCode.Text = Split(objEdit.Parent.Text, "-")(0)
            Me.txtCode.Text = Mid(Split(objEdit.Text, "-")(0), Len(Me.txtUpCode.Text) + 1)
            Me.txtCode.MaxLength = Len(Me.txtCode.Text)
            Me.txtCode.Tag = Me.txtCode.MaxLength
        End If
        Me.txtName.Text = Split(objEdit.Text, "-")(1)
        Me.txtExplain.Text = Split(objEdit.Tag, vbCrLf)(0)
        Me.lstScope.Tag = Split(objEdit.Tag, vbCrLf)(1) & "0000000"
        Me.lstScope.Selected(0) = (Mid(Me.lstScope.Tag, 1, 1) = "1")
        Me.lstScope.Selected(1) = (Mid(Me.lstScope.Tag, 2, 1) = "1")
        Me.lstScope.Selected(2) = (Mid(Me.lstScope.Tag, 3, 1) = "1")
        Me.lstScope.Selected(3) = (Mid(Me.lstScope.Tag, 4, 1) = "1")
        Me.lstScope.Selected(4) = (Mid(Me.lstScope.Tag, 5, 1) = "1")
        Me.lstScope.Selected(5) = (Mid(Me.lstScope.Tag, 6, 1) = "1")
        Me.lstScope.Selected(6) = (Mid(Me.lstScope.Tag, 7, 1) = "1")
        Me.lstScope.Selected(7) = (Mid(Me.lstScope.Tag, 8, 1) = "1")
        Me.Tag = Mid(objEdit.Key, 2)
    End If
    Me.Show vbModal, frmParent
    
    ShowMe = mblnOK
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Private Sub zlDefaultCode()
    '-----------------------------------------------------
    '���ܣ�����ѡ����ϼ�ID(�����Me.txtParent.Tag)���������ñ����ȱʡֵ
    '-----------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error Resume Next
    
    Me.chkCodeLen.Value = 0
    Me.chkCodeLen.Enabled = True
   
    If Me.txtParent.Tag = 0 Then
        Me.txtParent.Text = "(��)"
        Me.txtUpCode.Text = ""
        gstrSQL = "Select Max(����) As ���� From �����ʾ���� Where �ϼ�id Is Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        With rsTemp
            If IIf(IsNull(!����), "", !����) = "" Then
                Me.txtCode.Text = "01"
                Me.txtCode.MaxLength = mintMaxLen
                Me.txtCode.Tag = Me.txtCode.MaxLength
                Me.chkCodeLen.Value = 1
                Me.chkCodeLen.Enabled = False
            Else
                Me.txtCode.MaxLength = Len(Trim(!����))
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If !���� = String(Me.txtCode.MaxLength, "9") Then
                    If Me.txtCode.MaxLength >= mintMaxLen Then
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
            Me.txtUpCode.Text = Split(.SelectedItem.Text, "-")(0)
            If .SelectedItem.Children = 0 Then
                Me.txtCode.MaxLength = mintMaxLen - Len(Me.txtUpCode.Text)
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If Me.txtCode.MaxLength > 1 Then
                    Me.txtCode.Text = "01"
                Else
                    Me.txtCode.Text = "1"
                End If
                Me.chkCodeLen.Value = 1
                Me.chkCodeLen.Enabled = False
            Else
                gstrSQL = "Select Nvl(Max(����), '') As ���� From �����ʾ���� Where �ϼ�id = [1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(.SelectedItem.Key, 2)))
                With rsTemp
                    Me.txtCode.MaxLength = Len(!����) - Len(Me.txtUpCode.Text)
                    Me.txtCode.Tag = Me.txtCode.MaxLength
                    If Mid(!����, Len(Me.txtUpCode.Text) + 1) = String(Me.txtCode.MaxLength, "9") Then
                        If Len(Me.txtUpCode.Text) + Me.txtCode.MaxLength >= mintMaxLen Then
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


'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub chkCodeLen_Click()
    If Me.chkCodeLen.Value = 1 Then
        Me.txtCode.MaxLength = mintMaxLen - Len(Me.txtUpCode.Text)
    Else
        Me.txtCode.MaxLength = Me.txtCode.Tag
        Me.txtCode.Text = Mid(Me.txtCode.Text, 1, Me.txtCode.MaxLength)
    End If
    If Me.Visible Then Me.txtCode.SetFocus
End Sub

Private Sub chkCodeLen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
Dim lngItemID As Long
    
    Err = 0: On Error GoTo ErrHand
    
    Me.txtExplain.Text = Replace(Me.txtExplain.Text, vbCrLf, "")
    Me.txtExplain.Text = Replace(Me.txtExplain.Text, vbCr, "")
    Me.txtExplain.Text = Replace(Me.txtExplain.Text, vbLf, "")
    
    If Me.txtCode.MaxLength = 0 Then
        MsgBox "�ϼ������Ѿ��ﵽ��󳤶ȣ����������¼���", vbExclamation, gstrSysName
        Me.cmdCancel.SetFocus: Exit Sub
    End If
    If Trim(Me.txtCode.Text) = "" Then
        MsgBox "����������룡", vbExclamation, gstrSysName
        Me.txtCode.SetFocus: Exit Sub
    End If
    If Me.chkCodeLen.Value = 0 And Len(Trim(Me.txtCode.Text)) <> Me.txtCode.MaxLength Then
        MsgBox "���볤�ȱ���Ϊ" & Me.txtCode.MaxLength & "λ��������ѡ����ĳ���ѡ��", vbExclamation, gstrSysName
        Me.txtCode.SetFocus: Exit Sub
    End If
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "���Ʊ������룡", vbExclamation, gstrSysName
        Me.txtName.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "���Ƴ���" & Me.txtName.MaxLength & "�ĳ�������", vbExclamation, gstrSysName
        Me.txtName.SetFocus: Exit Sub
    End If
        
    Dim strApply As String
    strApply = ""
    For lngCount = 0 To Me.lstScope.ListCount - 1
        strApply = strApply & IIf(Me.lstScope.Selected(lngCount), "1", "0")
    Next
    If mblnAdd Then
        lngItemID = zlDatabase.GetNextId("�����ʾ����")
        gstrSQL = "Zl_�����ʾ����_Edit(1," & lngItemID & "," & Val(Me.txtParent.Tag) & ",'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "'," & _
            " '" & Trim(Me.txtName.Text) & "','" & Trim(Me.txtExplain.Text) & "','" & strApply & "'," & Me.chkCodeLen.Value & ")"
    Else
        lngItemID = Me.Tag
        gstrSQL = "Zl_�����ʾ����_Edit(2," & lngItemID & "," & Val(Me.txtParent.Tag) & ",'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "'," & _
            " '" & Trim(Me.txtName.Text) & "','" & Trim(Me.txtExplain.Text) & "','" & strApply & "'," & Me.chkCodeLen.Value & ")"
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnOK = True: Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelect_Click()
    With Me.tvwClass
        .Left = Me.txtParent.Left
        .Top = Me.txtParent.Top + Me.txtParent.Height
        .ZOrder 0: .Visible = True: .SetFocus
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.tvwClass.Visible Then
        Me.tvwClass.Visible = False: Me.txtParent.SetFocus: Exit Sub
    End If
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
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

Private Sub txtCode_GotFocus()
    Me.txtCode.SelStart = 0: Me.txtCode.SelLength = 100
    Call zlCommFun.OpenIme(False)
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

Private Sub txtExplain_GotFocus()
    Me.txtExplain.SelStart = 0: Me.txtExplain.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtExplain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    If InStr("' ", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    If InStr("' ", Chr(KeyAscii)) > 0 Then KeyAscii = 0
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

Private Sub txtUpCode_Change()
    Me.txtCode.Width = txtUpCode.Width - TextWidth(txtUpCode.Text) - 120
    Me.txtCode.Left = txtUpCode.Left + TextWidth(txtUpCode.Text) + 60
End Sub

