VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeSort 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�շ���Ŀ����༭"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   2310
      Left            =   -3810
      TabIndex        =   15
      Tag             =   "1000"
      Top             =   2010
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
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
   Begin VB.TextBox txtParent 
      Height          =   300
      Left            =   1725
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "(��)"
      ToolTipText     =   "��Del����ϼ������ó�������"
      Top             =   795
      Width           =   3495
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&P"
      Height          =   300
      Left            =   5235
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   795
      Width           =   285
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1725
      MaxLength       =   40
      TabIndex        =   8
      Tag             =   "����"
      Top             =   1815
      Width           =   3795
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4485
      TabIndex        =   14
      Top             =   3060
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3255
      TabIndex        =   13
      Top             =   3060
      Width           =   1100
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   -120
      TabIndex        =   12
      Top             =   2880
      Width           =   6345
   End
   Begin VB.CheckBox chkCodeLen 
      Caption         =   "������ı��볤�ȣ������˵�����ͬ������(&L)"
      Height          =   285
      Left            =   1005
      TabIndex        =   11
      Top             =   2595
      Width           =   4290
   End
   Begin VB.TextBox txtSymbol 
      Height          =   300
      Left            =   1725
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "����"
      Top             =   2175
      Width           =   1425
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1860
      MaxLength       =   15
      TabIndex        =   17
      Tag             =   "����"
      Text            =   "000000"
      Top             =   1455
      Width           =   1380
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3045
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   45
      Top             =   1185
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
            Picture         =   "frmChargeSort.frx":0000
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSort.frx":059A
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtUpCode 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1725
      MaxLength       =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "����"
      Text            =   "0000"
      Top             =   1410
      Width           =   1620
   End
   Begin VB.Label lblMsg 
      Caption         =   "���ӳɹ���"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblHint 
      Caption         =   "(��ʾ����Del����ϼ������ó�������)"
      Height          =   210
      Left            =   1695
      TabIndex        =   18
      Top             =   1155
      Width           =   3330
   End
   Begin VB.Label lblParent 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�(&U)"
      Height          =   180
      Left            =   1005
      TabIndex        =   2
      Top             =   855
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   1005
      TabIndex        =   7
      Top             =   1860
      Width           =   630
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "����(&D)"
      Height          =   180
      Left            =   1005
      TabIndex        =   5
      Top             =   1470
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmChargeSort.frx":0B34
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "    �շ���Ŀ�ɸ����ٴ���ҽ����������Ӧ�ò������ص����ͳһ�������ã��������������ӷ��ࡣ"
      Height          =   450
      Left            =   975
      TabIndex        =   1
      Top             =   165
      Width           =   4755
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ����"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   570
      Width           =   720
   End
   Begin VB.Label lblSymbol 
      AutoSize        =   -1  'True
      Caption         =   "����(&S)"
      Height          =   180
      Left            =   1005
      TabIndex        =   9
      Top             =   2235
      Width           =   630
   End
End
Attribute VB_Name = "frmChargeSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim intMaxLen As Integer
Dim objNode As Node
Dim mblnChanged As Boolean
Public mblnCancel As Boolean
Private mblnOk As Boolean

Public Function ShowMe(ByVal lngModle As Long, ByVal frmMain As Object) As Boolean
    Me.Show lngModle, frmMain
    ShowMe = mblnOk
End Function

Private Sub chkCodeLen_Click()
    On Error GoTo ErrHandle
    If Me.chkCodeLen.Value = 1 Then
        Me.txtCode.MaxLength = intMaxLen - Len(Me.txtUpCode.Text)
    Else
        Me.txtCode.MaxLength = Me.txtCode.Tag
        Me.txtCode.Text = Mid(Me.txtCode.Text, 1, Me.txtCode.MaxLength)
    End If
    Me.txtCode.SetFocus
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkCodeLen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
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
        lngItemID = sys.NextId("�շѷ���Ŀ¼")
        gstrSQL = "ZL_�շѷ���Ŀ¼_INSERT(" & _
            lngItemID & "," & _
            Me.txtParent.Tag & "," & _
            "'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "'," & _
            "'" & Trim(Me.txtName.Text) & "'," & _
            "'" & Trim(Me.txtSymbol.Text) & "'," & _
            Me.chkCodeLen.Value & ")"
    Else
        lngItemID = Me.Tag
        gstrSQL = "ZL_�շѷ���Ŀ¼_UPDATE(" & _
            lngItemID & "," & _
            Me.txtParent.Tag & "," & _
            "'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "'," & _
            "'" & Trim(Me.txtName.Text) & "'," & _
            "'" & Trim(Me.txtSymbol.Text) & "'," & _
            Me.chkCodeLen.Value & ")"
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mblnChanged = False
    mblnOk = True
    If Me.Tag = "����" Then
        txtName.Text = ""
        txtSymbol.Text = ""
        Call Form_Activate
        lblMsg.Visible = True
        txtName.SetFocus
        Exit Sub
    End If
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
        .ZOrder
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    Err = 0: On Error GoTo ErrHand
    
    If Me.Tag = "����" Then
        gstrSQL = "select ID,�ϼ�ID,����,����,����" & _
        " From �շѷ���Ŀ¼" & _
            " " & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    Else
        gstrSQL = "select ID,�ϼ�ID,����,����,����" & _
        " From �շѷ���Ŀ¼" & _
            " where id not in (select id from �շѷ���Ŀ¼ start with ID = [1] connect by prior id=�ϼ�id ) " & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.Tag))
    
    With rsTemp
        intMaxLen = .Fields("����").DefinedSize
        Me.txtName.MaxLength = .Fields("����").DefinedSize
        Me.txtSymbol.MaxLength = .Fields("����").DefinedSize
        
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�id) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�id, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    If Me.Tag = "����" Then
        mblnCancel = True
        Call zlDefaultCode
        mblnCancel = False
        txtName.SetFocus
    End If
    Me.txtCode.ZOrder
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mblnChanged = False
    mblnOk = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChanged = True And (mblnOk = False Or (mblnOk And txtSymbol.Text <> "" And txtName.Text <> "")) Then
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
    lblMsg.Visible = False
End Sub

Private Sub txtCode_GotFocus()
    Me.txtCode.SelStart = 0: Me.txtCode.SelLength = 100
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtName_Change()
    If mblnCancel = True Then Exit Sub
    mblnChanged = True
    lblMsg.Visible = False
End Sub

Private Sub txtName_GotFocus()
    Call OS.OpenIme(True)
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=`;'"":/.,?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.txtSymbol.Text = Mid(zlStr.GetCodeByORCL(Me.txtName.Text), 1, 10)
End Sub

Private Sub txtName_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txtParent_Change()
    If mblnCancel = True Then Exit Sub
    mblnChanged = True
    lblMsg.Visible = False
End Sub

Private Sub txtParent_GotFocus()
    Me.txtParent.SelStart = 0: Me.txtParent.SelLength = 100
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
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
        Call OS.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtSymbol_Change()
    If mblnCancel = True Then Exit Sub
    mblnChanged = True
    lblMsg.Visible = False
End Sub

Private Sub txtSymbol_GotFocus()
    Me.txtSymbol.SelStart = 0: Me.txtSymbol.SelLength = 100
End Sub

Private Sub txtSymbol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
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
    lblMsg.Visible = True
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
        gstrSQL = "select max(����) as ���� From �շѷ���Ŀ¼ Where �ϼ�ID is null "
        
    '            If .State = adStateOpen Then .Close
    '            Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "zlDefaultCode")
    '            Call SQLTest
        With rsTemp
            If IIF(IsNull(!����), "", !����) = "" Then
                Me.txtCode.Text = "01"
                Me.txtCode.MaxLength = intMaxLen
                Me.txtCode.Tag = Me.txtCode.MaxLength
                Me.chkCodeLen.Value = 1
                Me.chkCodeLen.Enabled = False
            Else
                Me.txtCode.MaxLength = Len(Trim(Nvl(!����)))
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If Nvl(!����) = String(Me.txtCode.MaxLength, "9") Then
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
                    Me.txtCode.Text = Format(Mid(Nvl(!����), Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                End If
            End If
        End With
    Else
        With Me.tvwClass
            .Nodes("_" & Me.txtParent.Tag).Selected = True
            Me.txtParent.Text = .SelectedItem.Text
            Me.txtUpCode.Text = Mid(Split(.SelectedItem.Text, "]")(0), 2)
            If .SelectedItem.Children = 0 Then
                Me.txtCode.MaxLength = IIF(intMaxLen - Len(Me.txtUpCode.Text) > 0, intMaxLen - Len(Me.txtUpCode.Text), 1)
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If Me.txtCode.MaxLength > 1 Then
                    Me.txtCode.Text = "01"
                Else
                    Me.txtCode.Text = "1"
                End If
                Me.chkCodeLen.Value = 1
                Me.chkCodeLen.Enabled = False
            Else
                gstrSQL = "select nvl(max(����),'') as ����  From �շѷ���Ŀ¼ Where �ϼ�ID=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(.SelectedItem.Key, 2)))
                                    
                With rsTemp
                    Me.txtCode.MaxLength = IIF(Len(Nvl(!����)) - Len(Me.txtUpCode.Text) > 0, Len(Nvl(!����)) - Len(Me.txtUpCode.Text), 1)
                    Me.txtCode.Tag = Me.txtCode.MaxLength
                    If Mid(Nvl(!����), Len(Me.txtUpCode.Text) + 1) = String(Me.txtCode.MaxLength, "9") Then
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
                        If Len(Nvl(!����)) >= Len(Me.txtUpCode.Text) + 1 Then
                            Me.txtCode.Text = Format(Mid(Nvl(!����), Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                        End If
                    End If
                End With
            End If
        End With
    End If
    Me.txtParent.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


