VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmProgPriv 
   Caption         =   "ģ��Ȩ������"
   ClientHeight    =   5565
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   8505
   Icon            =   "frmProgPriv.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8505
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   345
      Width           =   2025
   End
   Begin MSComctlLib.ListView lvwSelect 
      Height          =   1995
      Left            =   2985
      TabIndex        =   4
      Top             =   1860
      Visible         =   0   'False
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��������"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fra˵�� 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3000
      TabIndex        =   12
      Top             =   4485
      Width           =   5190
      Begin VB.TextBox txt˵�� 
         Height          =   300
         Left            =   1545
         MaxLength       =   125
         TabIndex        =   16
         Top             =   0
         Width           =   3630
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   420
         MaxLength       =   2
         TabIndex        =   14
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lbl˵�� 
         AutoSize        =   -1  'True
         Caption         =   "˵��"
         Height          =   180
         Left            =   1125
         TabIndex        =   15
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   0
         TabIndex        =   13
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "��ʵȨ��"
      Height          =   350
      Left            =   1470
      TabIndex        =   9
      Top             =   5070
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   3810
      Left            =   180
      TabIndex        =   1
      ToolTipText     =   "�޸�ģ��Ķ������Ȩ�޺�,�����ʹ�ø�ģ��Ľ�ɫ������Ȩ,Ȩ�޲�����Ч"
      Top             =   975
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   6720
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin ZL9BillEdit.BillEdit bill 
      Height          =   3195
      Left            =   2970
      TabIndex        =   5
      ToolTipText     =   "ɾ���尴Del��,�����������һ�е����һ�а��س�"
      Top             =   930
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   5636
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   4410
      TabIndex        =   8
      Top             =   5070
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   180
      TabIndex        =   10
      Top             =   5070
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5730
      TabIndex        =   6
      Top             =   5070
      Width           =   1100
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "��ԭ(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6960
      TabIndex        =   7
      Top             =   5070
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   1995
      Left            =   2910
      TabIndex        =   3
      Top             =   390
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   3519
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3915
      Left            =   2550
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3915
      ScaleWidth      =   45
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2460
      Top             =   2970
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
            Picture         =   "frmProgPriv.frx":0442
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProgPriv.frx":1296
            Key             =   "Dll"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ȩ��(&N)"
      Height          =   300
      Index           =   1
      Left            =   2880
      TabIndex        =   2
      Top             =   30
      Width           =   3285
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ģ��(&M)"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   675
      Width           =   1335
   End
End
Attribute VB_Name = "frmProgPriv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintSystem As Integer    'ϵͳ���
Private mstrOwner  As String     'ϵͳ������
Private mrsObject As New ADODB.Recordset

Private msngStartX As Single     '�ƶ�ǰ����λ��
Private mblnItem As Boolean
Private mstr��� As String        '��һ��tvwMain.SelectedItem��Key����
Private mstr���� As String    '��һ��Tab��Caption����

Private Sub bill_AfterDeleteRow()
    cmdRestore.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub bill_LostFocus()
    bill.TxtVisible = False
    bill.CmdVisible = False
End Sub

Private Sub bill_Validate(Cancel As Boolean)
If bill.TxtVisible = True And bill.Col = 0 Then
    If bill.Text = bill.TextMatrix(bill.Row, bill.Col) Then
       cmdSave.Enabled = False
       cmdRestore.Enabled = False
    End If
 End If
End Sub

Private Sub bill_CommandClick()
    If IsRecord("") = True Then
        lvwSelect.Top = bill.Top + bill.CellTop + bill.rowHeight(bill.Row)
        lvwSelect.Left = bill.Left + 15
        If bill.Top + bill.Height - lvwSelect.Top < lvwSelect.Height Then
            lvwSelect.Top = bill.Top + bill.CellTop - lvwSelect.Height
        End If
        lvwSelect.ZOrder 0
        lvwSelect.Visible = True
        lvwSelect.SetFocus
    End If
End Sub

Private Sub bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    
    If KeyCode = vbKeyReturn And bill.TxtVisible = True Then
        If IsRecord(bill.Text) = False Then
            Cancel = True
            bill.TxtSetFocus
        Else
            If lvwSelect.ListItems.Count = 1 Then
                If GetObject() = True Then
                    bill.Text = bill.TextMatrix(bill.Row, 0)
                Else
                    Cancel = True
                    bill.TxtSetFocus
                End If
            Else
                lvwSelect.Top = bill.Top + bill.CellTop + bill.rowHeight(bill.Row)
                lvwSelect.Left = bill.Left + 15
                lvwSelect.ZOrder 0
                lvwSelect.Visible = True
                lvwSelect.SetFocus
                Cancel = True
            End If
        End If
    ElseIf KeyCode = vbKeySpace And bill.Col > 1 Then
        If AllowSwitch() = False Then
            Cancel = True
            bill.TextMatrix(bill.Row, bill.Col) = " "
        Else
            cmdSave.Enabled = True
            cmdRestore.Enabled = True
        End If
    End If
End Sub

Private Sub bill_DblClick(Cancel As Boolean)

    If AllowSwitch() = False Then
        Cancel = True
        bill.TextMatrix(bill.Row, bill.Col) = " "
    Else
        If bill.Col > 1 Then
            cmdSave.Enabled = True
            cmdRestore.Enabled = True
        End If
    End If
End Sub

Private Function AllowSwitch() As Boolean
    Dim strType As String
    
    With bill
        strType = .TextMatrix(.Row, 1)
        Select Case .Col
            Case 0 '����
                AllowSwitch = True
                bill.TxtVisible = True
                bill.TxtEnable = True
            Case 2 'SELECT
                If strType = "SEQUENCE" Or strType = "TABLE" Or strType = "VIEW" Then AllowSwitch = True
            Case 3 'INSERT
                If strType = "TABLE" Or strType = "VIEW" Then AllowSwitch = True
            Case 4 'UPDATE
                If strType = "TABLE" Or strType = "VIEW" Then AllowSwitch = True
            Case 5 'DELETE
                If strType = "TABLE" Or strType = "VIEW" Then AllowSwitch = True
            Case 6 'EXECUTE
                If strType = "FUNCTION" Or strType = "PROCEDURE" Or strType = "PACKAGE" Or strType = "PACKAGE BODY" Or strType = "TYPE" Then AllowSwitch = True
        End Select
    End With
End Function

Private Sub bill_EnterCell(Row As Long, Col As Long)
'Ϊÿ�ֶ����ж�������
    With bill
        If .TextMatrix(Row, 0) <> "" And .TextMatrix(Row, 1) = "" Then
            mrsObject.Filter = "OBJECT_NAME='" & .TextMatrix(Row, 0) & "'"
            If Not mrsObject.EOF Then
                .TextMatrix(Row, 1) = mrsObject("OBJECT_TYPE")
            End If
        End If
    End With
End Sub

Private Sub cmbSystem_Click()
    mstrOwner = GetOwnerName(cmbSystem.ItemData(cmbSystem.ListIndex), gcnOracle)
    mintSystem = cmbSystem.ItemData(cmbSystem.ListIndex)
    gstrSQL = "Select Distinct Object_Name, Object_Type" & vbNewLine & _
            "From All_Objects" & vbNewLine & _
            "Where Owner = '" & mstrOwner & "' And Object_Type In ('FUNCTION', 'PACKAGE', 'PROCEDURE', 'TYPE', 'SEQUENCE', 'TABLE', 'VIEW') And" & vbNewLine & _
            "      Instr(Object_Name, 'BIN$') <= 0"

    Set mrsObject = gcnOracle.Execute(gstrSQL, adOpenStatic, adLockReadOnly)
    
    Call Fillģ��
    
    'ֻ�������߲�����ɱ�����
    cmdVerify.Enabled = (mstrOwner = gstrUserName)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.Name
End Sub

Private Sub cmdVerify_Click()
    Dim blnHaveErr As Boolean
    Dim blnStillHave As Boolean
    
    blnStillHave = Not frmModuleCheck.ShowMe(mintSystem, blnHaveErr)

    If Not blnHaveErr Then
        MsgBox "����飬Ȩ����ȫ��ȷ��", vbInformation, gstrSysName
    ElseIf Not blnStillHave Then
        MsgBox "������޸���Ȩ����ȫ��ȷ��", vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With cmbSystem
        .Top = ScaleTop + 30
        .Width = tvwMain.Width
        .Left = ScaleLeft
    End With
    
    lbl(0).Top = cmbSystem.Top + cmbSystem.Height + 30
    lbl(0).Width = tvwMain.Width
    tvwMain.Top = lbl(0).Top + lbl(0).Height + 30
    
    cmdHelp.Top = ScaleHeight - cmdHelp.Height - 100
    tvwMain.Height = IIf(cmdHelp.Top - 100 - tvwMain.Top > 0, cmdHelp.Top - 100 - tvwMain.Top, 0)
    tvwMain.Left = ScaleLeft
    lbl(0).Left = ScaleLeft
    
    picSplit.Top = ScaleTop
    picSplit.Height = tvwMain.Top + tvwMain.Height - picSplit.Top
    picSplit.Left = tvwMain.Left + tvwMain.Width
    
    lbl(1).Left = picSplit.Left + picSplit.Width
    lbl(1).Top = cmbSystem.Top
    tabMain.Top = tvwMain.Top
    tabMain.Height = tvwMain.Height - Me.fra˵��.Height - 45
    tabMain.Left = lbl(1).Left
    If ScaleWidth - tabMain.Left > 0 Then tabMain.Width = ScaleWidth - tabMain.Left
    lbl(1).Width = tabMain.Width
    
    bill.Left = tabMain.ClientLeft
    bill.Top = tabMain.ClientTop
    bill.Width = tabMain.ClientWidth
    bill.Height = tabMain.ClientHeight
    
    Me.fra˵��.Left = Me.tabMain.Left
    Me.fra˵��.Top = Me.tabMain.Top + Me.tabMain.Height + 45
    Me.fra˵��.Width = Me.tabMain.Width
    Me.txt˵��.Width = Me.fra˵��.Width - Me.txt˵��.Left - 45
    
    cmdClose.Top = cmdHelp.Top
    cmdSave.Top = cmdHelp.Top
    cmdRestore.Top = cmdHelp.Top
    cmdVerify.Top = cmdHelp.Top
    
    cmdRestore.Left = ScaleWidth - cmdRestore.Width - 200
    cmdSave.Left = cmdRestore.Left - cmdSave.Width - 150
    cmdClose.Left = cmdSave.Left - cmdClose.Width - 150
    Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdSave.Enabled = True Then
        If MsgBox("���ݶ���ķ���Ȩ���Ѹı䣬�Ƿ񱣴棿", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
            If SaveȨ�� = False Then
                MsgBox "�������", vbExclamation, gstrSysName
                Cancel = 1
            End If
        End If
    End If
End Sub

Private Sub lvwSelect_LostFocus()
    lvwSelect.Visible = False
End Sub

Private Sub lvwSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwSelect_ItemClick(ByVal Item As MSComctlLiB.ListItem)
    mblnItem = True
End Sub

Private Sub lvwSelect_DblClick()
    If mblnItem = True Then
        Call lvwSelect_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub lvwSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If GetObject() = True Then
            lvwSelect.Visible = False
            bill.SetFocus
            SendKeys "{ENTER}"
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        lvwSelect.Visible = False
        bill.SetFocus
        If bill.TxtVisible = True Then
            bill.TxtSetFocus
        End If
    End If
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - msngStartX
        If sngTemp > 500 And ScaleWidth - (sngTemp + picSplit.Width) > 1000 Then
            picSplit.Left = sngTemp
            tvwMain.Width = picSplit.Left - tvwMain.Left
            tabMain.Left = picSplit.Left + picSplit.Width
            tabMain.Width = ScaleWidth - tabMain.Left
            lbl(0).Width = tvwMain.Width
            lbl(1).Left = tabMain.Left
            lbl(1).Width = tabMain.Width
            bill.Left = tabMain.ClientLeft
            bill.Top = tabMain.ClientTop
            bill.Width = tabMain.ClientWidth
            bill.Height = tabMain.ClientHeight
        End If
    End If
End Sub

Private Sub tvwMain_NodeClick(ByVal Node As MSComctlLiB.Node)
    Call Fill����
End Sub

Private Sub TabMain_Click()
    Call Fill��Ȩ��
    bill.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call SaveȨ��
End Sub

Private Sub cmdRestore_Click()
    If MsgBox("���ݶ���ķ���Ȩ���Ѹı䣬�Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    mstr���� = ""
    cmdSave.Enabled = False
    Call Fill��Ȩ��
End Sub

Private Sub Fillģ��()
    Dim rsTemp As New ADODB.Recordset
    Dim str����  As String
    
    mstr��� = ""

    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Popedom.Get_Module", mintSystem)
    tvwMain.Nodes.Clear
    Do Until rsTemp.EOF
        If rsTemp("����") <> str���� Then
            str���� = rsTemp("����")
            tvwMain.Nodes.Add , , "C" & rsTemp("����"), rsTemp("����"), "Dll", "Dll"
        End If
        tvwMain.Nodes.Add "C" & rsTemp("����"), tvwChild, "C" & rsTemp("���"), rsTemp("����"), "Module", "Module"
        rsTemp.MoveNext
    Loop
    '���µ���
    tabMain.Enabled = rsTemp.RecordCount > 0
    bill.Enabled = tabMain.Enabled
    If tvwMain.Nodes.Count > 0 Then tvwMain.Nodes(1).Selected = True
    Call Fill����
End Sub

Private Sub Fill����()
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    Dim str���� As String
    
    
    If tvwMain.SelectedItem Is Nothing Then
        strKey = "0"
    Else
        If tvwMain.SelectedItem.Image = "Dll" Then
            strKey = Mid(tvwMain.SelectedItem.Child.Key, 2)
            str���� = tvwMain.SelectedItem.Child.Text
        Else
            strKey = Mid(tvwMain.SelectedItem.Key, 2)
            str���� = tvwMain.SelectedItem.Text
        End If
    End If
    If mstr��� = strKey Then Exit Sub
    If cmdSave.Enabled = True Then
        If MsgBox("���ݶ���ķ���Ȩ���Ѹı䣬�Ƿ񱣴棿", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
            Call SaveȨ��
        End If
        cmdSave.Enabled = False
        cmdRestore.Enabled = False
    End If
    mstr��� = strKey
    mstr���� = ""
    lbl(1).Caption = str���� & "ģ���Ȩ��(&N)"
    '��ҳ����и���
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Popedom.Get_function", mintSystem, Val(strKey))
    mstr���� = ""
    tabMain.Tabs.Clear
    tabMain.Tabs.Add , "C����", "����"
    Do Until rsTemp.EOF
        If rsTemp("����") <> "����" Then
            tabMain.Tabs.Add , "C" & rsTemp("����"), rsTemp("����")
        End If
        rsTemp.MoveNext
    Loop
    tabMain.Enabled = rsTemp.RecordCount > 0
    bill.Enabled = tabMain.Enabled
    Call Fill��Ȩ��
End Sub

Private Sub Fill��Ȩ��()
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    rsTemp.CursorLocation = adUseClient
    If tabMain.SelectedItem.Caption = mstr���� Then Exit Sub
    If cmdSave.Enabled = True Then
        If MsgBox("���ݶ���ķ���Ȩ���Ѹı䣬�Ƿ񱣴棿", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
            Call SaveȨ��
        End If
    End If
    mstr���� = tabMain.SelectedItem.Caption
    
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Popedom.Get_function", mintSystem, Val(mstr���), mstr����)
    If Not rsTemp.EOF Then
        Me.txt����.Text = IIf(IsNull(rsTemp!����), 0, rsTemp!����)
        Me.txt˵��.Text = IIf(IsNull(rsTemp!˵��), "", rsTemp!˵��)
        Me.txt����.Enabled = True: Me.txt˵��.Enabled = True
    Else
        Me.txt����.Text = ""
        Me.txt˵��.Text = ""
        Me.txt����.Enabled = False: Me.txt˵��.Enabled = False
    End If
    rsTemp.Close

    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Popedom.Get_impower", mintSystem, Val(mstr���), mstr����)
    If rsTemp.RecordCount = 0 Then
        bill.Rows = 2
        For i = 0 To bill.Cols - 1
            bill.TextMatrix(1, i) = ""
        Next
    Else
        bill.Rows = rsTemp.RecordCount + 1
    End If
    i = 1
    Do Until rsTemp.EOF
        bill.TextMatrix(i, 0) = rsTemp("����")
        bill.TextMatrix(i, 1) = ""
        bill.TextMatrix(i, 2) = IIf(rsTemp("SELECT") > 0, "��", " ")
        bill.TextMatrix(i, 3) = IIf(rsTemp("INSERT") > 0, "��", " ")
        bill.TextMatrix(i, 4) = IIf(rsTemp("UPDATE") > 0, "��", " ")
        bill.TextMatrix(i, 5) = IIf(rsTemp("DELETE") > 0, "��", " ")
        bill.TextMatrix(i, 6) = IIf(rsTemp("EXECUTE") > 0, "��", " ")
        rsTemp.MoveNext
        i = i + 1
    Loop
    cmdRestore.Enabled = False
    cmdSave.Enabled = False
End Sub

Private Function IsRecord(ByVal strWhere As String) As Boolean
    Dim strTemp As String
    
    IsRecord = False
    If InStr(strWhere, "'") > 0 Then
        MsgBox "�����˷Ƿ��ַ���'����", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If Trim(strWhere) = "" Then
        mrsObject.Filter = 0
    Else
        mrsObject.Filter = "OBJECT_NAME LIKE '" & Trim(strWhere) & "%'"
    End If
    If mrsObject.RecordCount = 0 Then
        MsgBox "û�к��ʵĶ���", vbExclamation, gstrSysName
        Exit Function
    End If
    lvwSelect.ListItems.Clear
    Do Until mrsObject.EOF
        lvwSelect.ListItems.Add , "C" & mrsObject("OBJECT_NAME"), mrsObject("OBJECT_NAME")
        lvwSelect.ListItems("C" & mrsObject("OBJECT_NAME")).SubItems(1) = mrsObject("OBJECT_TYPE")
        mrsObject.MoveNext
    Loop
    lvwSelect.ListItems(1).Selected = True
    IsRecord = True
End Function

Private Function GetObject() As Boolean
    Dim i As Integer
    
    For i = 1 To bill.Rows - 1
        If bill.TextMatrix(i, 0) = lvwSelect.SelectedItem.Text And i <> bill.Row Then
            MsgBox "�����ݶ����Ѿ�ѡ���ˡ�", vbExclamation, gstrSysName
            Exit Function
        End If
    Next
    With lvwSelect.SelectedItem
        bill.TextMatrix(bill.Row, 0) = .Text
        bill.TextMatrix(bill.Row, 1) = .SubItems(1)
        
        lvwSelect.Visible = False
        If .SubItems(1) = "TABLE" Or .SubItems(1) = "VIEW" Or .SubItems(1) = "SEQUENCE" Then
            bill.TextMatrix(bill.Row, 2) = "��"
            bill.TextMatrix(bill.Row, 6) = " "
        Else
            bill.TextMatrix(bill.Row, 6) = "��"
        End If
        bill.Col = 6
        If Trim(bill.Text) <> Trim(bill.TextMatrix(bill.Row, bill.Col)) Then
            cmdRestore.Enabled = True
            cmdSave.Enabled = True
        End If
    End With
    GetObject = True
End Function

Private Function SaveȨ��() As Boolean
    Dim intRow As Integer, intCol As Integer
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    gcnOracle.Execute "Update zlProgFuncs set ����=" & Val(Me.txt����.Text) & ",˵��='" & Trim(Me.txt˵��.Text) & "' where ϵͳ=" & mintSystem & " and ���=" & mstr��� & " and ����='" & mstr���� & "'"
    gcnOracle.Execute "delete from zlProgPrivs where ϵͳ=" & mintSystem & " and ���=" & mstr��� & " and ����='" & mstr���� & "'"
    For intRow = 1 To bill.Rows - 1
        If bill.TextMatrix(intRow, 0) <> "" Then
            For intCol = 2 To bill.Cols - 1
                If Trim(bill.TextMatrix(intRow, intCol)) <> "" Then
                    gstrSQL = "insert into zlProgPrivs (ϵͳ,���,����,����,������,Ȩ��) values " & _
                        "(" & mintSystem & "," & mstr��� & ",'" & mstr���� & "','" & UCase(bill.TextMatrix(intRow, 0)) & _
                        "','" & UCase(mstrOwner) & "','" & bill.TextMatrix(0, intCol) & "')"
                    gcnOracle.Execute gstrSQL
                End If
            Next
        End If
    Next
    '������Ҫ������־
    Call SaveAuditLog(1, "�޸�ģ���ʹ��Ȩ��", "�޸�ģ���ʹ��Ȩ��")
    gcnOracle.CommitTrans
    SaveȨ�� = True
    cmdSave.Enabled = False
    cmdRestore.Enabled = False
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    MsgBox "����ʧ�ܡ�" & vbNewLine & err.Description, vbExclamation, gstrSysName
End Function

Public Function ProgPriv() As Boolean
    '��ʼ�����
    With bill
        .Cols = 7
        .TextMatrix(0, 0) = "���ݶ���"
        .TextMatrix(0, 1) = "��������"
        .TextMatrix(0, 2) = "SELECT"
        .TextMatrix(0, 3) = "INSERT"
        .TextMatrix(0, 4) = "UPDATE"
        .TextMatrix(0, 5) = "DELETE"
        .TextMatrix(0, 6) = "EXECUTE"
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 0
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .ColAlignment(5) = 4
        .ColAlignment(6) = 4
    
        .ColData(0) = 1
        .ColData(1) = 5
        .ColData(2) = -1
        .ColData(3) = -1
        .ColData(4) = -1
        .ColData(5) = -1
        .ColData(6) = -1
        
        .PrimaryCol = 0
        .Active = True
    End With

    Call FillSystem

    '9i��֧��
    gstrSQL = "Select Distinct Object_Name, Object_Type" & vbNewLine & _
            "From All_Objects" & vbNewLine & _
            "Where Owner = '" & mstrOwner & "' And Object_Type In ('FUNCTION', 'PACKAGE', 'PROCEDURE', 'TYPE', 'SEQUENCE', 'TABLE', 'VIEW') And" & vbNewLine & _
            "      Instr(Object_Name, 'BIN$') <= 0"

    Set mrsObject = gcnOracle.Execute(gstrSQL, adOpenStatic, adLockReadOnly)
    
    Call Fillģ��
    
    'ֻ�������߲�����ɱ�����
    cmdVerify.Enabled = (mstrOwner = gstrUserName)
    
    '��ʾ����
    frmProgPriv.Show vbModal, frmMDIMain
    If mrsObject.State = 1 Then mrsObject.Close
    Set mrsObject = Nothing
End Function

Private Sub FillSystem()
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle

    '��ʾ�������е�ϵͳ
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    Set rsTemp = zlGetRegSystems
    cmbSystem.Clear
    Do Until rsTemp.EOF
        cmbSystem.addItem RPAD(rsTemp("����") & "��" & rsTemp("���") & "��", 25) & " v" & rsTemp("�汾��")
        cmbSystem.ItemData(cmbSystem.NewIndex) = rsTemp("���")
        If glngSysNo = rsTemp("���") And cmbSystem.ListIndex < 0 Then
            cmbSystem.ListIndex = cmbSystem.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If cmbSystem.ListIndex < 0 Then cmbSystem.ListIndex = 0

    '��ϵͳ��¼
    If glngSysNo <> -1 Then
        cmbSystem.Enabled = False
        '��ȡϵͳ������
        mstrOwner = GetOwnerName(glngSysNo, gcnOracle)
        mintSystem = glngSysNo
    Else
        mstrOwner = GetOwnerName(cmbSystem.ItemData(0), gcnOracle)
        mintSystem = cmbSystem.ItemData(0)
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub txt����_Change()
    cmdSave.Enabled = True
    cmdRestore.Enabled = True
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = Me.txt����.MaxLength
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Me.txt˵��.SetFocus
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt˵��_Change()
    cmdSave.Enabled = True
    cmdRestore.Enabled = True
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = Me.txt˵��.MaxLength
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Me.cmdSave.Enabled = True Then
        Me.cmdSave.SetFocus
    End If
End Sub
