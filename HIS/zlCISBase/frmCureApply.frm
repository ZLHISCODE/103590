VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmCureApply 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ο����÷�Χ"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "frmCureApply.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�(&R)"
      Height          =   350
      Left            =   2520
      Picture         =   "frmCureApply.frx":058A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4110
      Width           =   1170
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ�����(&E)"
      Height          =   350
      Left            =   1350
      Picture         =   "frmCureApply.frx":06D4
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4110
      Width           =   1170
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5505
      Top             =   4560
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
            Picture         =   "frmCureApply.frx":081E
            Key             =   "ItemUse"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3825
      TabIndex        =   3
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   225
      Picture         =   "frmCureApply.frx":0DB8
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4920
      TabIndex        =   4
      Top             =   4110
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit msfApply 
      Height          =   2895
      Left            =   225
      TabIndex        =   2
      Top             =   1095
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
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
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2700
      Left            =   510
      TabIndex        =   6
      Top             =   4470
      Visible         =   0   'False
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   4763
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
   Begin VB.Label lblRefer 
      AutoSize        =   -1  'True
      Caption         =   "�ο���Ŀ��"
      Height          =   180
      Left            =   780
      TabIndex        =   1
      Top             =   795
      Width           =   900
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCureApply.frx":0F02
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   5220
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   225
      Picture         =   "frmCureApply.frx":0F94
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frmCureApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1��ָ����Ŀ����me.lblRefer.tag���棬���ϼ�����ͨ��ShowMe�������룬���Դ��ݣ�Ҳ���Բ�����
'---------------------------------------------------
Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer

Public Sub ShowME(ByVal frmParent As Object, ByVal lng��Ŀid As Long)
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    Me.msfApply.Active = True
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.����,I.����,I.���� from ���Ʋο�Ŀ¼ I Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng��Ŀid)
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ���ǰ�Ĳο�(���ܱ�������ɾ��)��", vbExclamation, gstrSysName
            Exit Sub
        Else
            Me.lblRefer.Tag = !ID: Me.lblRefer.Caption = "�ο���Ŀ��[" & !���� & "]" & !����
            Me.Tag = !����
            Call zlApplyItem(Me.lblRefer.Tag)
        End If
    End With
    Me.Show 1, frmParent
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClear_Click()
    Me.msfApply.ClearBill
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call zlApplyItem(Me.lblRefer.Tag)
End Sub

Private Sub cmdOk_Click()
    strTemp = "": gstrSql = ""
    With Me.msfApply
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" And .RowData(intCount) <> 0 Then
                If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "����Ŀ��ǰ����Ŀ���ظ���", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
                strTemp = strTemp & ";" & .RowData(intCount)
                gstrSql = gstrSql & "|" & .RowData(intCount)
            End If
        Next
    End With
    If gstrSql <> "" Then gstrSql = Mid(gstrSql, 2)
    gstrSql = "zl_���Ʋο���Χ_Save(" & Val(Me.lblRefer.Tag) & ",'" & gstrSql & "')"
    
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False
        Me.msfApply.SetFocus
    Else
        cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msfApply
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 2
        .TextMatrix(0, 0) = "���": .TextMatrix(0, 1) = "������Ŀ"
        .ColData(0) = 5: .ColData(1) = 1
        .ColWidth(0) = 500: .ColWidth(1) = 5200
        .ColAlignment(0) = 4
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 3300
        .Add , "����", "����", 1200
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
        Me.msfApply.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
        Me.msfApply.RowData(Me.msfApply.Row) = Mid(.SelectedItem.Key, 2)
        Me.msfApply.TextMatrix(Me.msfApply.Row, 1) = Me.msfApply.Text
        Me.msfApply.SetFocus
        Call zlCommFun.PressKey(vbKeyRight)
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

Private Sub msfApply_AfterAddRow(Row As Long)
    With Me.msfApply
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfApply_AfterDeleteRow()
    With Me.msfApply
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfApply_CommandClick()
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "Select I.ID, I.����, I.����" & _
            "   From ������ĿĿ¼ I, ���Ʒ���Ŀ¼ C" & _
            "  Where I.����id = C.Id And C.���� =[1] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by I.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "Ŀǰû�н���ͬ����Ŀ���޷����ã�", vbExclamation, gstrSysName: Exit Sub
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
        .Tag = Me.msfApply.Name
        .Left = Me.msfApply.Left + 300
        .Top = Me.msfApply.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfApply_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfApply.TextMatrix(Row, Col)
End Sub

Private Sub msfApply_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfApply_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msfApply
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
    
    gstrSql = "Select distinct I.ID,I.����,I.����,I.�걾��λ" & _
            " From ������ĿĿ¼ I,������Ŀ���� N,���Ʒ���Ŀ¼ C" & _
            " Where I.ID=N.������Ŀid And I.����id = C.Id And C.����=[1] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.���� like [2] or N.���� like [3] or N.���� like [3])" & _
            " Order by I.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.Tag), strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "δָ����ͬ����Ŀ�����������룡", vbExclamation, gstrSysName: Cancel = True: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfApply.Text = "[" & !���� & "]" & !����
            Me.msfApply.RowData(Me.msfApply.Row) = !ID
            Me.msfApply.TextMatrix(Me.msfApply.Row, 1) = Me.msfApply.Text
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
        .Tag = Me.msfApply.Name
        .Left = Me.msfApply.Left + 300
        .Top = Me.msfApply.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlApplyItem(lngItemId As Long)
    '--------------------------------------------------------
    '���ܣ�ˢ����ʾ���õ�������Ŀ
    '��Σ�lngItemId-ָ���Ĳο���Ŀid
    '--------------------------------------------------------
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,'['||I.����||']'||I.���� as ����" & _
            " from ������ĿĿ¼ I" & _
            " where I.�ο�Ŀ¼ID=[1] " & _
            " Order by I.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    
    With rsTemp
        Me.msfApply.ClearBill
        Do While Not .EOF
            If Me.msfApply.Rows - 1 < .AbsolutePosition Then Me.msfApply.Rows = Me.msfApply.Rows + 1
            Me.msfApply.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfApply.RowData(.AbsolutePosition) = !ID
            Me.msfApply.TextMatrix(.AbsolutePosition, 1) = !����
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


