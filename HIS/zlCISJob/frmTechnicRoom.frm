VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTechnicRoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ִ�м�����"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "frmTechnicRoom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   30
      TabIndex        =   10
      Top             =   555
      Width           =   2865
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   3210
      Picture         =   "frmTechnicRoom.frx":058A
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Left            =   3210
      Picture         =   "frmTechnicRoom.frx":06D4
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1335
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   795
      MaxLength       =   20
      TabIndex        =   4
      Top             =   3780
      Width           =   3525
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   3210
      TabIndex        =   5
      Top             =   2730
      Width           =   1100
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�(&R)"
      Height          =   350
      Left            =   3210
      TabIndex        =   6
      Top             =   3120
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwRoom 
      Height          =   2730
      Left            =   105
      TabIndex        =   0
      Top             =   945
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   4815
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&X)"
      Height          =   350
      Left            =   3210
      TabIndex        =   7
      Top             =   120
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3375
      Top             =   3765
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
            Picture         =   "frmTechnicRoom.frx":081E
            Key             =   "Room"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNote 
      Caption         =   "    ���ñ����ҵ�ִ�м�󣬲�����Ч����ִ�еİ��š�"
      Height          =   405
      Left            =   150
      TabIndex        =   9
      Top             =   120
      Width           =   2700
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "��ҽѧ��ִ�м�:"
      Height          =   180
      Left            =   150
      TabIndex        =   8
      Top             =   720
      Width           =   1350
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   105
      TabIndex        =   3
      Top             =   3825
      Width           =   630
   End
End
Attribute VB_Name = "frmTechnicRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem

Private Sub cmdAdd_Click()
    Me.lblName.Tag = "": Me.txtName.Text = "": Me.txtName.Enabled = True
    Me.cmdDel.Enabled = True: Me.cmdSave.Enabled = True: Me.cmdRestore.Enabled = True
    Me.txtName.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me: Exit Sub
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    
    If MsgBox("���ɾ��ִ�м䡰" & Trim(Me.txtName.Text) & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    strSQL = "zl_ҽ��ִ�з���_Delete(" & Val(Me.lblDept.Tag) & ",'" & Trim(Me.txtName.Text) & "')"
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Call zlRoomRef
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRestore_Click()
    Call zlRoomRef
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String
    
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
    '-----------------------------------------
    Err = 0: On Error GoTo ErrHand
    If Me.lblName.Tag = "" Then
        strSQL = "zl_ҽ��ִ�з���_Insert(" & Val(Me.lblDept.Tag) & ",'" & Trim(Me.txtName.Text) & "')"
    Else
        strSQL = "zl_ҽ��ִ�з���_Update(" & Val(Me.lblDept.Tag) & ",'" & Trim(Me.lblName.Tag) & "','" & Trim(Me.txtName.Text) & "')"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

    MsgBox "ִ�м䱣��ɹ���", vbExclamation, gstrSysName
    
    Call zlRoomRef
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub Form_Activate()
    Call zlRoomRef
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    ElseIf InStr("|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Me.lvwRoom.ListItems.Clear
    With Me.lvwRoom.ColumnHeaders
        .Clear
        .Add , "����", "����", 2500
    End With
End Sub

Private Sub lvwRoom_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtName.Text = Item.Text
    Me.lblName.Tag = Me.txtName.Text
End Sub

Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub zlRoomRef()
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    strSQL = "Select ִ�м� From ҽ��ִ�з��� where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Me.lblDept.Tag))
    Me.lvwRoom.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwRoom.ListItems.Add(, , !ִ�м�, 1, 1)
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
    If Me.lvwRoom.ListItems.Count > 0 Then
        Me.lvwRoom.ListItems(1).Selected = True
        Me.lvwRoom.SelectedItem.EnsureVisible
    End If
    
    Err = 0: On Error GoTo 0
    If Me.lvwRoom.ListItems.Count > 0 Then
        Call lvwRoom_ItemClick(Me.lvwRoom.SelectedItem)
        Me.txtName.Enabled = True
        Me.cmdDel.Enabled = True: Me.cmdSave.Enabled = True: Me.cmdRestore.Enabled = True
    Else
        Me.lblName.Tag = "": Me.txtName.Text = ""
        Me.txtName.Enabled = False
        Me.cmdDel.Enabled = False: Me.cmdSave.Enabled = False: Me.cmdRestore.Enabled = False
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsTemp = Nothing
    Set objItem = Nothing
End Sub