VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectColumn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ѡ����"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   Icon            =   "frmSelectColumn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�Ĭ������(&R)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3150
      TabIndex        =   12
      Top             =   3990
      Width           =   1695
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "����(&D)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3750
      TabIndex        =   11
      Top             =   3450
      Width           =   1100
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "����(&U)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3750
      TabIndex        =   10
      Top             =   3030
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&L)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3750
      TabIndex        =   9
      Top             =   2610
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫѡ(&S)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3750
      TabIndex        =   8
      Top             =   2190
      Width           =   1100
   End
   Begin VB.ComboBox cmbAlign 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4380
      Width           =   1185
   End
   Begin VB.TextBox txtWidth 
      Height          =   300
      Left            =   1290
      TabIndex        =   2
      Top             =   3960
      Width           =   1155
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3750
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3750
      TabIndex        =   6
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3750
      TabIndex        =   5
      Top             =   240
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwColumn 
      Height          =   3645
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   6429
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��ע"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label lblWidth 
      Caption         =   "�п�(&W)"
      Height          =   180
      Left            =   540
      TabIndex        =   1
      Top             =   4020
      Width           =   630
   End
   Begin VB.Label lblAlign 
      Caption         =   "���뷽ʽ(&A)"
      Height          =   180
      Left            =   180
      TabIndex        =   3
      Top             =   4440
      Width           =   990
   End
End
Attribute VB_Name = "frmSelectColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mstrKey As String        '��ǰ
Dim mstrType As String       '�ؼ�����
Dim mstrColumns As String    '�ؼ���ȱʡ������
Dim mobjSet As Object        'Ҫ���õĿؼ�

Private Sub cmbAlign_Click()
    Select Case mstrType
        Case "ListView"
            If cmbAlign.Text = "���ж���" Then
                lvwColumn.SelectedItem.ListSubItems(1).Tag = 2
            ElseIf cmbAlign.Text = "���Ҷ���" Then
                lvwColumn.SelectedItem.ListSubItems(1).Tag = 1
            Else
                lvwColumn.SelectedItem.ListSubItems(1).Tag = 0
            End If
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    mblnOK = False
End Sub

Private Sub cmdOK_Click()
    Select Case mstrType
        Case "ListView"
            SetLvwColumns mobjSet
    End Select
    Unload Me
    mblnOK = True
End Sub

Private Sub SetLvwColumns(lvwTemp As Object)
    Dim str���� As String
    Dim lngCount As Long
    Dim lst As Object
    Dim i As Integer
    Dim intPos As Integer
    
    str���� = lvwColumn.FindItem("����", lvwSubItem).Text
    LockWindowUpdate lvwTemp.hwnd
    lvwTemp.ColumnHeaders.Clear
    lvwTemp.ColumnHeaders.Add , "_" & str����, str����
    For lngCount = 1 To lvwColumn.ListItems.Count
        Set lst = lvwColumn.ListItems(lngCount)
        If lst.Checked = True Then
            i = i + 1 '����λ��
            If lst.Text <> str���� Then
                lvwTemp.ColumnHeaders.Add , "_" & lst.Text, lst.Text, lst.Tag, lst.ListSubItems(1).Tag
            Else
                intPos = i
                lvwTemp.ColumnHeaders(1).Width = lst.Tag
                lvwTemp.ColumnHeaders(1).Alignment = lst.ListSubItems(1).Tag
            End If
        End If
    Next
    lvwTemp.ColumnHeaders(1).Position = intPos
    
    LockWindowUpdate 0
End Sub

Public Function ������(objSet As Object, ByVal strColumn As String) As Boolean
    Dim varColumns As Variant
    Dim varColumn As Variant
    Dim lngCol As Long
    Dim lst As ListItem
    
    mstrColumns = strColumn
    mstrType = TypeName(objSet)
    Set mobjSet = objSet
    varColumns = Split(strColumn, ";")
    lvwColumn.ListItems.Clear
    For lngCol = LBound(varColumns) To UBound(varColumns)
        varColumn = Split(varColumns(lngCol), ",")
        Set lst = lvwColumn.ListItems.Add(, varColumn(0), varColumn(0))
        lst.Tag = varColumn(1) '�����п��
        Select Case mstrType
            Case "ListView"
                 Select Case varColumn(3)
                    Case 1
                        lst.SubItems(1) = "����"
                    Case 2
                        lst.SubItems(1) = "��������"
                    Case Else
                        lst.SubItems(1) = ""
                End Select
            Case "MSHFlexGrid"
            Case "DataGrid"
        End Select
        lst.ListSubItems(1).Tag = varColumn(2) '���뷽ʽ
    Next
    cmbAlign.AddItem "���ж���"
    cmbAlign.AddItem "�������"
    cmbAlign.AddItem "���Ҷ���"
    '�ѿؼ������ݽ������
    Select Case mstrType
        Case "ListView"
            Dim colTemp As ColumnHeader
            For Each colTemp In objSet.ColumnHeaders
                Set lst = lvwColumn.ListItems(colTemp.Text)
                lst.Tag = Round(colTemp.Width) '�����п��
                lst.ListSubItems(1).Tag = colTemp.Alignment '���뷽ʽ
                lst.Checked = True
                UpListView lst.Key, colTemp.Position
            Next
        Case "MSHFlexGrid"
        Case "DataGrid"
    End Select
    
    lvwColumn.ListItems(1).Selected = True
    lvwColumn_ItemClick lvwColumn.ListItems(1)
    frmSelectColumn.Show vbModal
    ������ = mblnOK
End Function

Private Sub cmdRestore_Click()
    Dim varColumns As Variant
    Dim varColumn As Variant
    Dim lngCol As Long
    Dim lst As ListItem
    
    
    LockWindowUpdate lvwColumn.hwnd
    varColumns = Split(mstrColumns, ";")
    lvwColumn.ListItems.Clear
    For lngCol = LBound(varColumns) To UBound(varColumns)
        varColumn = Split(varColumns(lngCol), ",")
        Set lst = lvwColumn.ListItems.Add(, varColumn(0), varColumn(0))
        lst.Tag = varColumn(1) '�����п��
        Select Case mstrType
            Case "ListView"
                 Select Case varColumn(3)
                    Case 1
                        lst.SubItems(1) = "����"
                    Case 2
                        lst.SubItems(1) = "��������"
                    Case Else
                        lst.SubItems(1) = ""
                End Select
            Case "MSHFlexGrid"
            Case "DataGrid"
        End Select
        lst.ListSubItems(1).Tag = varColumn(2) '���뷽ʽ
        lst.Checked = True
    Next
    LockWindowUpdate 0
    mstrKey = ""
    lvwColumn.ListItems(1).Selected = True
    lvwColumn_ItemClick lvwColumn.ListItems(1)
End Sub

Private Sub cmdSelect_Click()
    Dim lngCount As Long
    
    For lngCount = 1 To lvwColumn.ListItems.Count
        lvwColumn.ListItems(lngCount).Checked = True
    Next
End Sub

Private Sub cmdClear_Click()
    Dim lngCount As Long
    
    For lngCount = 1 To lvwColumn.ListItems.Count
        If lvwColumn.ListItems(lngCount).SubItems(1) = "" Then lvwColumn.ListItems(lngCount).Checked = False
    Next
End Sub

Private Sub cmdUp_Click()
    Dim strKey As String
    
    If lvwColumn.SelectedItem Is Nothing Then Exit Sub
    If lvwColumn.SelectedItem.Index = 1 Then Exit Sub
                
    LockWindowUpdate lvwColumn.hwnd
    strKey = lvwColumn.SelectedItem.Key
    UpListView strKey, lvwColumn.SelectedItem.Index - 1
    LockWindowUpdate 0
    lvwColumn.ListItems(strKey).Selected = True
End Sub

Private Sub UpListView(ByVal strKey As String, ByVal lngNewIndex As Long)
    Dim varTemp(1 To 5) As Variant
    With lvwColumn.ListItems(strKey)
        varTemp(1) = .Tag
        varTemp(2) = .ListSubItems(1).Text
        varTemp(3) = .ListSubItems(1).Tag
        varTemp(4) = .Index
        varTemp(5) = .Checked
    End With
    lvwColumn.ListItems.Remove strKey
    lvwColumn.ListItems.Add lngNewIndex, strKey, strKey
    lvwColumn.ListItems(strKey).Tag = varTemp(1)
    lvwColumn.ListItems(strKey).SubItems(1) = varTemp(2)
    lvwColumn.ListItems(strKey).ListSubItems(1).Tag = varTemp(3)
    lvwColumn.ListItems(strKey).Checked = varTemp(5)
End Sub

Private Sub cmdDown_Click()
    Dim strKey As String
    
    If lvwColumn.SelectedItem Is Nothing Then Exit Sub
    If lvwColumn.SelectedItem.Index = lvwColumn.ListItems.Count Then Exit Sub
                
    LockWindowUpdate lvwColumn.hwnd
    strKey = lvwColumn.SelectedItem.Key
    UpListView strKey, lvwColumn.SelectedItem.Index + 1
    LockWindowUpdate 0
    lvwColumn.ListItems(strKey).Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnOK = False
    mstrKey = ""
    mstrType = ""
End Sub

Private Sub lvwColumn_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.SubItems(1) <> "" Then Item.Checked = True
    If Item Is lvwColumn.SelectedItem Then
        SetEnable Item.Checked
        If Item.SubItems(1) = "����" Then cmbAlign.Enabled = False
    End If
End Sub

Private Sub lvwColumn_ItemClick(ByVal Item As MSComctlLib.ListItem)
'������ʾ���⣬��Ҫ�����п�Ͷ��뷽ʽ
    Dim itmTemp As MSComctlLib.ListItem
    
    If mstrKey = Item.Key Then Exit Sub
    
    If mstrKey <> "" Then
        Set itmTemp = lvwColumn.ListItems(mstrKey)
        itmTemp.Tag = Val(txtWidth)
        
        Select Case mstrType
            Case "ListView"
                If cmbAlign.Text = "���ж���" Then
                    itmTemp.ListSubItems(1).Tag = 2
                ElseIf cmbAlign.Text = "���Ҷ���" Then
                    itmTemp.ListSubItems(1).Tag = 1
                Else
                    itmTemp.ListSubItems(1).Tag = 0
                End If
        End Select
    End If
    mstrKey = Item.Key
    
    txtWidth.Text = Item.Tag
    Select Case mstrType
        Case "ListView"
            If Item.ListSubItems(1).Tag = 2 Then
                cmbAlign.Text = "���ж���"
            ElseIf Item.ListSubItems(1).Tag = 1 Then
                cmbAlign.Text = "���Ҷ���"
            Else
                cmbAlign.Text = "�������"
            End If
    End Select
    SetEnable Item.Checked
    If mstrType = "ListView" And Item.SubItems(1) = "����" Then cmbAlign.Enabled = False
End Sub

Private Sub txtWidth_Change()
    If Trim(txtWidth) <> "" Then Call IsValid
End Sub

Private Function IsValid() As Boolean
    Dim blnValid As Boolean
    
    blnValid = True
    If IsNumeric(txtWidth) = False Then
        blnValid = False
        MsgBox "������һ���Ϸ�����ֵ��", vbInformation, gstrSysName
    Else
        If Val(txtWidth.Text) > 10000 Or Val(txtWidth.Text) < 0 Then
            MsgBox "������һ��С��10000��������", vbInformation, gstrSysName
            blnValid = False
        End If
    End If
    IsValid = blnValid
End Function

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyReturn And _
        KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> Asc(".") Then KeyAscii = 0
End Sub

Private Sub txtWidth_Validate(Cancel As Boolean)
    Cancel = Not IsValid
    If Cancel = False Then
        lvwColumn.SelectedItem.Tag = Val(txtWidth)
    End If
End Sub
Private Sub SetEnable(ByVal blnEnable As Boolean)
    lblAlign.Enabled = blnEnable
    lblWidth.Enabled = blnEnable
    txtWidth.Enabled = blnEnable
    cmbAlign.Enabled = blnEnable
End Sub
