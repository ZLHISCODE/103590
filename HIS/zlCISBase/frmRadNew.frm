VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRadNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ӱ����Ŀ����"
   ClientHeight    =   6360
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8400
   Icon            =   "frmRadNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8400
   StartUpPosition =   1  '����������
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   405
      Left            =   6840
      TabIndex        =   20
      Top             =   3840
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   714
      ButtonWidth     =   1349
      ButtonHeight    =   609
      TextAlignment   =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ȫѡ"
            Key             =   "ȫѡ"
            Object.ToolTipText     =   "ѡ��������ʾ��Ŀ"
            Object.Tag             =   "ȫѡ"
            ImageKey        =   "SelectAll"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ȫ��"
            Key             =   "ȫ��"
            Object.ToolTipText     =   "�������ѡ���־"
            Object.Tag             =   "ȫ��"
            ImageKey        =   "ClearAll"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ӱ���鲹����Ϣ"
      Height          =   1575
      Left            =   2760
      TabIndex        =   19
      Top             =   4200
      Width           =   5655
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   555
         Width           =   2055
      End
      Begin VB.ComboBox cbo��Ƭ 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   870
         Width           =   2055
      End
      Begin VB.TextBox txt׼�� 
         Height          =   300
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   14
         Top             =   1185
         Width           =   4230
      End
      Begin VB.TextBox txtͼ�� 
         Height          =   300
         Left            =   3855
         MaxLength       =   2
         TabIndex        =   12
         Top             =   870
         Width           =   780
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         Caption         =   "Ӱ�����"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "���в���"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl��Ƭ 
         AutoSize        =   -1  'True
         Caption         =   "�ɷ���Ƭ"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   930
         Width           =   720
      End
      Begin VB.Label lbl׼�� 
         AutoSize        =   -1  'True
         Caption         =   "���׼��"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   1245
         Width           =   720
      End
      Begin VB.Label lblͼ�� 
         AutoSize        =   -1  'True
         Caption         =   "�������ͼ����Ŀ"
         Height          =   180
         Left            =   3840
         TabIndex        =   11
         Top             =   630
         Width           =   1440
      End
   End
   Begin VB.CheckBox chkOnly 
      Caption         =   "ֻ��ʾ�����Ŀ(&C)"
      Height          =   255
      Left            =   2910
      TabIndex        =   4
      Top             =   3915
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&X)"
      Height          =   350
      Left            =   7110
      TabIndex        =   18
      Top             =   5940
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   195
      Picture         =   "frmRadNew.frx":058A
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5940
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   5970
      TabIndex        =   16
      Top             =   5940
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   15
      Top             =   5820
      Width           =   8535
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -30
      TabIndex        =   1
      Top             =   510
      Width           =   8535
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   60
      Top             =   4785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadNew.frx":06D4
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadNew.frx":0C6E
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadNew.frx":1208
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadNew.frx":1422
            Key             =   "ClearAll"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   5190
      Left            =   0
      TabIndex        =   2
      Tag             =   "1000"
      Top             =   585
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   9155
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   3255
      Left            =   2760
      TabIndex        =   3
      Top             =   570
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   60
      Picture         =   "frmRadNew.frx":163C
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    Ӱ������Ŀֻ�ܴ��Ѿ����������ü����������Ŀ��ѡ�����ӣ�Ȼ�󲹳��Ҫ��Ӱ������Ϣ���Ӷ���֤���ٴ�Ӧ�õ�һ���ԡ�"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   630
      TabIndex        =   0
      Top             =   90
      Width           =   7650
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuSele 
         Caption         =   "ȫ��ѡ��(&A)"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPopuSele 
         Caption         =   "ȫ��ȡ��(&R)"
         Index           =   1
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmRadNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem

Dim strTemp As String, aryTemp() As String
Dim intCount As Integer

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo��Ƭ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkOnly_Click()
    If Me.Tag = "Loading" Then Exit Sub
    LoadClass
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strDescribe As String
    
    strDescribe = "'" & Split(Me.cbo���.Text, "-")(0) & "'"
    strDescribe = strDescribe & "," & Left(Me.cbo����.Text, 1)
    strDescribe = strDescribe & "," & Left(Me.cbo��Ƭ.Text, 1)
    strDescribe = strDescribe & ",'" & Trim(Me.txt׼��.Text) & "'"
    strDescribe = strDescribe & "," & Val(Me.txtͼ��.Text)
    
    For Each objItem In Me.lvwItem.ListItems
        If objItem.Checked Then
            gstrSql = "zl_Ӱ������Ŀ_Insert(" & Mid(objItem.Key, 2) & "," & strDescribe & ")"
            Err = 0: On Error Resume Next
            Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
            If Err <> 0 Then
                Call SaveErrLog
            End If
        End If
    Next
    
    MsgBox "�������õ�Ӱ������Ŀ������ϣ�", vbExclamation, gstrSysName
    Call tvwClass_NodeClick(Me.tvwClass.SelectedItem)
    Call frmRadLists.zlRefItems
End Sub

Private Sub cmd����_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Me.lvwItem.ColumnHeaders.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_����", "����", 1000
        .Add , "_����", "����", 2500
        .Add , "_��λ", "��λ", 1000
        .Add , "_���㵥λ", "��λ", 600
    End With
    With Me.lvwItem
        .ColumnHeaders("_����").Position = 1
        .SortKey = .ColumnHeaders("_����").Index - 1: .SortOrder = lvwAscending
    End With
    
    '---------------------------------
    'װ���������
    gstrSql = "Select * From Ӱ������� Order By ����"
    Err = 0: On Error GoTo ErrHand
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    Me.cbo���.Clear
    With rsTemp
        Do While Not .EOF
            Me.cbo���.AddItem !���� & "-" & !����
            If !���� = Mid(frmRadLists.lvwKind.SelectedItem.Key, 2) Then
                Me.cbo���.ListIndex = Me.cbo���.NewIndex
            End If
            .MoveNext
        Loop
    End With
        
    aryTemp = Split("0-������;1-����;2-ѡ�����", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo����.AddItem aryTemp(intCount)
    Next
    Me.cbo����.ListIndex = 0
    
    aryTemp = Split("0-������;1-����;2-ѡ�񷢷�", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo��Ƭ.AddItem aryTemp(intCount)
    Next
    Me.cbo��Ƭ.ListIndex = 0
    
    Me.Tag = "Loading"
    chkOnly.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "Ӱ��ֻѡ������Ŀ", 1))
    Me.Tag = ""
    LoadClass
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadClass()
    Dim strCurrKey As String
    '----------------------------------
'    gstrSql = "Select Distinct ID, ����, ����, �ϼ�ID" & _
'            " From ���Ʒ���Ŀ¼" & _
'            " Where ���� = '5'" & _
'            " Start With id In (Select Distinct ����id From ������ĿĿ¼ Where ��� = 'D' Or" & _
'            " (���='E' And ��������='5') Or ���='Z')" & _
'            " Connect By Prior �ϼ�ID = ID" & _
'            " Order By ����"
    gstrSql = "Select Distinct ID, ����, ����, �ϼ�ID" & _
            " From ���Ʒ���Ŀ¼" & _
            " Where ���� = '5'" & _
            " Start With id In (Select Distinct ����id From ������ĿĿ¼" & _
            IIf(chkOnly.Value = 1, " Where ��� = 'D')", ")") & _
            " Connect By Prior �ϼ�ID = ID" & _
            " Order By ����"
    
    Err = 0: On Error GoTo ErrHand
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "LoadClass")
'        Call SQLTest
    With rsTemp
        If Not tvwClass.SelectedItem Is Nothing Then strCurrKey = tvwClass.SelectedItem.Key
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            If strCurrKey = objNode.Key Then objNode.Selected = True
            .MoveNext
        Loop
    End With
    Err = 0: On Error GoTo 0
    If Me.tvwClass.Nodes.count > 0 Then
        If tvwClass.SelectedItem Is Nothing Then Me.tvwClass.Nodes(1).Selected = True
        tvwClass.SelectedItem.EnsureVisible
        Call tvwClass_NodeClick(Me.tvwClass.SelectedItem)
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "Ӱ��ֻѡ������Ŀ", chkOnly.Value)
End Sub

Private Sub lvwItem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.mnuPopu, 2
    End If
End Sub

Private Sub mnuPopuSele_Click(Index As Integer)
    For Each objItem In Me.lvwItem.ListItems
        objItem.Checked = (Index = 0)
    Next
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
     Select Case Button.Key
        Case "ȫѡ"
            SelectAll True
        Case "ȫ��"
            SelectAll False
     End Select
End Sub

Private Sub SelectAll(ByVal blnSelect As Boolean)
    Dim i As Integer
    With lvwItem
        For i = 1 To .ListItems.count
            .ListItems(i).Checked = blnSelect
        Next
    End With
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    
    gstrSql = "Select I.ID,I.����, I.����,I.�걾��λ, I.���㵥λ" & _
            "   From ������ĿĿ¼ I" & _
            " Where " & IIf(chkOnly.Value = 1, "��� = 'D' And ", "") & "����id In " & _
            "       (Select id From ���Ʒ���Ŀ¼ Start With id = [1] Connect By Prior id = �ϼ�ID)" & _
            "       And (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            " Minus" & _
            " Select I.ID,I.����, I.����,I.�걾��λ, I.���㵥λ" & _
            "   From ������ĿĿ¼ I, Ӱ������Ŀ R" & _
            "  Where I.ID = R.������Ŀid"
    
    Err = 0: On Error GoTo ErrHand
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(Node.Key, 2))
        
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_��λ").Index - 1) = IIf(IsNull(!�걾��λ), "", !�걾��λ)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_���㵥λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            .MoveNext
        Loop
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtͼ��_GotFocus()
    Me.txtͼ��.SelStart = 0: Me.txtͼ��.SelLength = 100
End Sub

Private Sub txtͼ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt׼��_GotFocus()
    Me.txt׼��.SelStart = 0: Me.txt׼��.SelLength = Me.txt׼��.MaxLength
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt׼��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt׼��_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub
