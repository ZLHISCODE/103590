VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSrcEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Դ�༭"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "frmSrcEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6345
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdData 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   2740
      TabIndex        =   5
      ToolTipText     =   "�������ݵ��뷽ʽ"
      Top             =   5070
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѡ��ODBC����Դ"
      Height          =   3135
      Left            =   720
      TabIndex        =   15
      Top             =   1680
      Width           =   5415
      Begin MSComctlLib.ListView lvwDSN 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   1960
         TabIndex        =   9
         ToolTipText     =   "����ODBC����Դ"
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   4200
         TabIndex        =   11
         ToolTipText     =   "ɾ��ODBC����Դ"
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdModi 
         Caption         =   "�޸�(&M)"
         Height          =   350
         Left            =   3075
         TabIndex        =   10
         ToolTipText     =   "����ODBC����Դ"
         Top             =   200
         Width           =   1100
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3855
      TabIndex        =   6
      Top             =   5070
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   270
      Picture         =   "frmSrcEdit.frx":058A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5070
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4980
      TabIndex        =   7
      Top             =   5070
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   4920
      Width           =   6315
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Top             =   585
      Width           =   5505
   End
   Begin VB.TextBox txt˵�� 
      Height          =   555
      Left            =   1425
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1110
      Width           =   4575
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1425
      MaxLength       =   50
      TabIndex        =   1
      Top             =   720
      Width           =   4590
   End
   Begin VB.Label lblNote 
      Caption         =   "ϵͳͨ������ODBC����Դ��ʵ�ָ���ҽ�����ݵĵ��룬�û����ڴ��趨ϵͳ��ҽ�������ļ������Ӻ͵��뷽ʽ��"
      Height          =   345
      Left            =   735
      TabIndex        =   13
      Top             =   120
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmSrcEdit.frx":06D4
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbl˵�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "˵��(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   2
      Top             =   1125
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   780
      Width           =   630
   End
End
Attribute VB_Name = "frmSrcEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSourceSQL As String, strDestFields As String, ifDeleteData As Boolean
Private strNewName As String

Private ifOK As Boolean
Private OldSourceName As String, OldDSN As String
Public Function EditSource(ByVal frmParent As Object, ByRef SourceName As String) As Boolean
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    On Error Resume Next
    EditSource = False
    OldSourceName = SourceName
    strSourceSQL = "": strDestFields = "": ifDeleteData = False: OldDSN = "": strNewName = ""
    
    If Len(SourceName) > 0 Then
        Me.txt���� = SourceName
        Me.txt˵�� = GetSetting("ZLSOFT", "ҽ������\" & SourceName, "˵��", "")
        OldDSN = GetSetting("ZLSOFT", "ҽ������\" & SourceName, "ODBC", "")
        
        strSourceSQL = GetSetting("ZLSOFT", "ҽ������\" & SourceName, "����Դ", "")
        strDestFields = GetSetting("ZLSOFT", "ҽ������\" & SourceName, "�ֶ�", "")
        ifDeleteData = GetSetting("ZLSOFT", "ҽ������\" & SourceName, "�������", "false")
    End If
    
    ListSource True
    '��ʾ����
    Me.Show 1, frmParent
    EditSource = ifOK: If EditSource Then SourceName = strNewName
End Function

Private Sub cmdAdd_Click()
    Dim curIndex As Long
    If CreateDataSource(Me.hWnd) Then
        curIndex = lvwDSN.SelectedItem.Index
        Call ListSource: lvwDSN.SelectedItem = lvwDSN.ListItems(curIndex): lvwDSN.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdDel_Click()
    Dim curIndex As Long
    If lvwDSN.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("�Ƿ�ɾ������Դ��" + lvwDSN.SelectedItem.Text + "��", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    If RemoveDataSource(Me.hWnd, lvwDSN.SelectedItem.Text, lvwDSN.SelectedItem.SubItems(1)) Then
        curIndex = lvwDSN.SelectedItem.Index
        Call ListSource
        
        On Error Resume Next
        If curIndex > lvwDSN.ListItems.Count - 1 Then curIndex = curIndex - 1
        If curIndex > -1 Then lvwDSN.SelectedItem = lvwDSN.ListItems(curIndex)
        lvwDSN.SetFocus
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    If Trim(Me.txt����.Text) = "" Then MsgBox "���������ƣ�", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���" & CInt(Me.txt����.MaxLength / 2) & "�����֣���", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    End If
    If Me.lvwDSN.SelectedItem Is Nothing Then
        MsgBox "��ѡ��һ��ODBC����Դ��", vbInformation, gstrSysName: Exit Sub
    End If
    If Len(Trim(strSourceSQL)) = 0 Then
        If MsgBox("δ�������ݵ��뷽ʽ���Ƿ������", vbDefaultButton2 + vbYesNo + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    '�ж������Ƿ��ظ�
    If UCase(Trim(OldSourceName)) <> UCase(Trim(Me.txt����)) And _
        GetSetting("ZLSOFT", "ҽ������", UCase(Trim(Me.txt����)), Chr(0)) <> Chr(0) Then _
            MsgBox "����Դ��" & Me.txt���� & "�Ѵ��ڣ�����������", vbInformation, gstrSysName: _
            Me.txt����.SetFocus: Exit Sub
    
    If Len(OldSourceName) > 0 Then
        Call DeleteSetting("ZLSOFT", "ҽ������", UCase(OldSourceName))
        Call DeleteSetting("ZLSOFT", "ҽ������\" & UCase(OldSourceName))
    End If
    Call SaveSetting("ZLSOFT", "ҽ������", UCase(Trim(Me.txt����)), "1")
    Call SaveSetting("ZLSOFT", "ҽ������\" & UCase(Trim(Me.txt����)), "˵��", Me.txt˵��)
    Call SaveSetting("ZLSOFT", "ҽ������\" & UCase(Trim(Me.txt����)), "ODBC", Me.lvwDSN.SelectedItem.Text)
    Call SaveSetting("ZLSOFT", "ҽ������\" & UCase(Trim(Me.txt����)), "����Դ", strSourceSQL)
    Call SaveSetting("ZLSOFT", "ҽ������\" & UCase(Trim(Me.txt����)), "�ֶ�", strDestFields)
    Call SaveSetting("ZLSOFT", "ҽ������\" & UCase(Trim(Me.txt����)), "�������", CStr(ifDeleteData))
    
    ifOK = True: strNewName = UCase(Me.txt����)
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdModi_Click()
    Dim curIndex As Long
    If lvwDSN.SelectedItem Is Nothing Then Exit Sub
    
    If ConfigDataSource(Me.hWnd, lvwDSN.SelectedItem.Text, lvwDSN.SelectedItem.SubItems(1)) Then
        curIndex = lvwDSN.SelectedItem.Index
        Call ListSource: lvwDSN.SelectedItem = lvwDSN.ListItems(curIndex): lvwDSN.SetFocus
    End If
End Sub

Private Sub cmdData_Click()
    Dim strSQL As String, strFlds As String, ifClear As Boolean
    
    If lvwDSN.SelectedItem Is Nothing Then Exit Sub
    
    frmDataSet.ShowMe Me, lvwDSN.SelectedItem.Text, strSQL, strFlds, ifClear
    If Len(Trim(strSQL)) > 0 Then
        strSourceSQL = strSQL: strDestFields = strFlds: ifDeleteData = ifClear
    End If
End Sub

Private Sub Form_Activate()
    '��ȡִ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_Load()
    ifOK = False
    With lvwDSN.ColumnHeaders
        .Add , "_Name", "����Դ", 1800
        .Add , "_Desc", "˵��", 3000
    End With
    lvwDSN.Sorted = True
End Sub

Private Sub ListSource(Optional ByVal ifInit As Boolean = False)
    Dim strDrivers As String, aDrivers() As String
    Dim i As Integer, tmpItem As ListItem, aSourceInfo() As String
    lvwDSN.ListItems.Clear
    
    strDrivers = GetODBCSources
    If Len(strDrivers) > 0 Then
        aDrivers = Split(strDrivers, Chr(0) + Chr(0))

        For i = 0 To UBound(aDrivers, 1)
            aSourceInfo = Split(aDrivers(i), Chr(0))
            Set tmpItem = lvwDSN.ListItems.Add(, "_" & i, aSourceInfo(0))
            tmpItem.SubItems(1) = aSourceInfo(1)
            
            If ifInit And UCase(aSourceInfo(0)) = UCase(OldDSN) Then tmpItem.Selected = True
        Next
    End If
End Sub

Private Sub lvwDSN_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwDSN
        .SortKey = ColumnHeader.Index - 1: .SortOrder = (.SortOrder + 1) Mod 2: .Sorted = True
    End With
End Sub

Private Sub lvwDSN_DblClick()
    cmdModi_Click
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_LostFocus()
    Me.txt˵��.Text = Replace(Me.txt˵��, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub


