VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmCureRBans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����֢����"
   ClientHeight    =   3810
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6480
   Icon            =   "frmClinicBans.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   0
      TabIndex        =   7
      Top             =   690
      Width           =   7530
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ�����(&R)"
      Height          =   350
      Left            =   1575
      Picture         =   "frmClinicBans.frx":058A
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3345
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4005
      TabIndex        =   2
      Top             =   3345
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5100
      TabIndex        =   3
      Top             =   3345
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3345
      Width           =   1290
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2715
      Left            =   390
      TabIndex        =   6
      Top             =   3885
      Visible         =   0   'False
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   5745
      Top             =   3960
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
            Picture         =   "frmClinicBans.frx":06D4
            Key             =   "ItemUse"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfBans 
      Height          =   2355
      Left            =   285
      TabIndex        =   1
      Top             =   855
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   4154
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
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   255
      Picture         =   "frmClinicBans.frx":0C6E
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ��ȷȫ������ö�Ӧ����Ժ;��Խ���֢(Ŀǰ������ҽ����)�����԰���ҽ���˽���Ŀ��Ӧ�ã�����ҽ�����"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   855
      TabIndex        =   0
      Top             =   165
      Width           =   5310
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "���ѡ��"
      Visible         =   0   'False
      Begin VB.Menu mnuKind 
         Caption         =   "����ҩ(&1)"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmCureRBans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strBans As String

Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer
Dim aryTemp() As String, strTemp As String

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdClear_Click()
    Me.msfBans.ClearBill
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    strBans = "": strTemp = ""
    With Me.msfBans
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" And .RowData(intCount) <> 0 Then
                If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "�У��ظ�ѡ���ˡ�" & .TextMatrix(intCount, 1) & "����", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
                strTemp = strTemp & ";" & .RowData(intCount)
                strBans = strBans & "|" & .RowData(intCount) & "^" & IIf(Trim(.TextMatrix(intCount, 2)) = "���", 1, 2)
            End If
        Next
    End With
    If strBans = "" Then
        If MsgBox("û��ѡ���ȫ����������ƴ�ʩ��������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    If strBans <> "" Then strBans = Mid(strBans, 2)
    Me.Hide
End Sub

Private Sub Form_Activate()
    If strBans = "" Then Exit Sub
    
    'װ���Ѿ�ѡ��Ľ���֢
    aryTemp = Split(strBans, "|")
    Err = 0: On Error GoTo ErrHand
    strTemp = ""
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        strTemp = strTemp & "," & Left(aryTemp(intCount), InStr(1, aryTemp(intCount), "^") - 1)
    Next
    
    gstrSql = "select I.ID,I.����,I.����" & _
            " from �������Ŀ¼ I" & _
            " where I.���=1 and I.ID in ([1])" & _
            " order by I.����"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(strTemp, 2))
    
    With rsTemp
        Me.msfBans.ClearBill
        Do While Not .EOF
            If Me.msfBans.Rows - 1 < .AbsolutePosition Then Me.msfBans.Rows = Me.msfBans.Rows + 1
            Me.msfBans.RowData(.AbsolutePosition) = !ID
            Me.msfBans.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfBans.TextMatrix(.AbsolutePosition, 1) = "[" & !���� & "]" & !����
            For intCount = LBound(aryTemp) To UBound(aryTemp)
                If Val(Left(aryTemp(intCount), InStr(1, aryTemp(intCount), "^") - 1)) = !ID Then
                    If Val(Mid(aryTemp(intCount), InStr(1, aryTemp(intCount), "^") + 1)) = 1 Then
                        Me.msfBans.TextMatrix(.AbsolutePosition, 2) = "���"
                    Else
                        Me.msfBans.TextMatrix(.AbsolutePosition, 2) = "����"
                    End If
                    Exit For
                End If
            Next
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False
        Me.msfBans.SetFocus
    Else
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msfBans
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 3
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "��������": .TextMatrix(0, 2) = "��������"
        .ColData(0) = 5: .ColData(1) = 1: .ColData(2) = 3
        .ColWidth(0) = 250: .ColWidth(1) = 4000: .ColWidth(2) = 1000
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
        .AddItem "���": .ItemData(.NewIndex) = 1
        .AddItem "����": .ItemData(.NewIndex) = 2
        .ListIndex = 0
    End With
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 3800
        .Add , "����", "����", 1000
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
    End With
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        Me.msfBans.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
        Me.msfBans.RowData(Me.msfBans.Row) = Mid(.SelectedItem.Key, 2)
        Me.msfBans.TextMatrix(Me.msfBans.Row, 1) = Me.msfBans.Text
        Me.msfBans.SetFocus
        Call zlCommFun.PressKey(vbKeyReturn)
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msfBans_AfterAddRow(Row As Long)
    With Me.msfBans
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfBans_AfterDeleteRow()
    With Me.msfBans
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfBans_cboClick(ListIndex As Long)
    Me.msfBans.TextMatrix(Me.msfBans.Row, 2) = Me.msfBans.CboText
End Sub

Private Sub msfBans_CommandClick()
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.����,I.���� from �������Ŀ¼ I where I.���=1 and (I.����ʱ�� Is Null Or I.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) order by I.����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "cmd����_Click")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "Ŀǰû�н����������Ŀ¼���޷����ã�", vbExclamation, gstrSysName: Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Left = Me.msfBans.Left + 300
        .Top = Me.msfBans.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfBans_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfBans.TextMatrix(Row, Col)
End Sub

Private Sub msfBans_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfBans_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msfBans
        If .Active = False Then Exit Sub
        If .Col <> 1 Then Exit Sub
        If .TxtVisible = False Then
            If .TextMatrix(.Row, 1) = "" Then Exit Sub
            strTemp = UCase(Trim(.TextMatrix(.Row, 1)))
        Else
            If Trim(.Text) = "" Then Exit Sub
            strTemp = UCase(Trim(.Text))
        End If
    End With
    If strTemp = strInputed Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.����,I.����" & _
            " from �������Ŀ¼ I,������ϱ��� N" & _
            " where I.ID=N.���id and I.���=1" & _
            "       and (I.���� like [1] or N.���� like [2] or N.���� like [2])" & _
            " and (I.����ʱ�� Is Null Or I.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ���ļ�����ϣ����������룡", vbExclamation, gstrSysName: Cancel = True: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfBans.Text = "[" & !���� & "]" & !����
            Me.msfBans.RowData(Me.msfBans.Row) = !ID
            Me.msfBans.TextMatrix(Me.msfBans.Row, 1) = Me.msfBans.Text
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.msfBans.Name
        .Left = Me.msfBans.Left + 300
        .Top = Me.msfBans.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
